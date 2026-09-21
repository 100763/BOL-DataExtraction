using iQPUBLIC;
using System;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace BOLSET2IQ
{
    /// <summary>
    /// clsParseUSAddress => Static class containing method to parse US address and get different sections
    /// </summary>
    public static class clsParseUSAddress
    {
        #region Not in Use - SQL Connection Passed
        ///// <summary>
        ///// Method to parse the US address and split into different sections - Street#, Direction, AddrLine1, City, State, ZipCode
        ///// </summary>
        ///// <param name="AddressLineTextValue">Entire US Address Line Text</param>
        ///// <param name="SQLConnectionObject">SQL Connection object created using DB connection string. Connection object is opened before use and closed after use in this method.</param>
        ///// <returns>AddressSection class object containing 6 attributes - StreetNumber, Direction, AddressLine1, City, State, ZipCode</returns>
        //public static AddressSection ParseUSAddress(string AddressLineTextValue, SqlConnection SQLConnectionObject)
        //{
        //    AddressSection objAddressSections = new AddressSection();
        //    try
        //    {
        //        string AddressLineText = AddressLineTextValue;

        //        // Check if address line contains "PO Box" keywords => "PO Box", "P.O. Box", P.O.Box", remove all PO Box keywords and the immediate number
        //        bool isPOBOXPresent = false;
        //        bool isSuitePresent = false;
        //        if (AddressLineText.ToUpper().Contains("PO BOX") || AddressLineText.ToUpper().Contains("P.O. BOX") ||
        //            AddressLineText.ToUpper().Contains("P.O.BOX") || AddressLineText.ToUpper().Contains("P O BOX"))
        //        {
        //            isPOBOXPresent = true;
        //        }

        //        // Check if address line contains "Suite" keywords => "Suite", "STE" remove all Suite keywords and the immediate number
        //        if (AddressLineText.ToUpper().Contains("SUITE") || AddressLineText.ToUpper().Contains("STE"))
        //        {
        //            isSuitePresent = true;
        //        }

        //        // Step 1 - Split the address line text into separate words using "single space"
        //        string[] splittedArray = AddressLineText.Split(' ');
        //        string[] newSplittedArray = new string[splittedArray.Length + 1];

        //        // Remove PO BOX, Suite keywords and immediate next word from address
        //        if (isPOBOXPresent || isSuitePresent)
        //        {
        //            if (isPOBOXPresent)
        //            {
        //                int j = 0;
        //                for (int index = 0; index < splittedArray.Length; index++)
        //                {
        //                    string word = splittedArray[index];
        //                    if (word.ToUpper() == "PO" || word.ToUpper() == "P.O.")
        //                    {
        //                        index = index + 2;
        //                    }
        //                    else if (word.ToUpper() == "P.O.BOX")
        //                    {
        //                        index = index + 1;
        //                    }
        //                    else if (word.ToUpper() == "P")
        //                    {
        //                        if (index + 1 < splittedArray.Length)
        //                        {
        //                            if (splittedArray[index + 1].ToUpper() == "O")
        //                            {
        //                                index = index + 3;
        //                            }
        //                        }
        //                    }
        //                    else
        //                    {
        //                        newSplittedArray[j] = word;
        //                        j++;
        //                    }
        //                }
        //            }
        //            else if (isSuitePresent)
        //            {
        //                int j = 0;
        //                for (int index = 0; index < splittedArray.Length; index++)
        //                {
        //                    string word = splittedArray[index];
        //                    if (word.ToUpper() == "SUITE" || word.ToUpper() == "STE")
        //                    {
        //                        index = index + 1;
        //                    }
        //                    else
        //                    {
        //                        newSplittedArray[j] = word;
        //                        j++;
        //                    }
        //                }
        //            }
        //        }
        //        else
        //        {
        //            newSplittedArray = splittedArray;
        //        }

        //        newSplittedArray = newSplittedArray.Where(item => item != null).ToArray();

        //        if (newSplittedArray.Length > 0)
        //        {
        //            // Check if the last word of the array is only numeric zip code or contains combined state code with zip code
        //            // if yes, split the state code and zip code into 2 separate words - state code and zip code
        //            string[] finalSplittedArray = new string[newSplittedArray.Length + 1];

        //            string lastWord = newSplittedArray[newSplittedArray.Length - 1];
        //            Match matchNew = Regex.Match(lastWord, "^[0-9-]+$");
        //            if (matchNew.Success == false)
        //            {
        //                int digitCount = lastWord.Count(item => Char.IsDigit(item));
        //                if (digitCount == 5 || digitCount == 9 || digitCount == 3 || digitCount == 4) // Zip code could be of 3, 4, 5 or 9 digits length
        //                {
        //                    for (int rowIndex = 0; rowIndex < newSplittedArray.Length; rowIndex++)
        //                    {
        //                        finalSplittedArray[rowIndex] = newSplittedArray[rowIndex];
        //                    }

        //                    if (lastWord.Contains(","))
        //                    {
        //                        string[] stateZip = lastWord.Split(',');
        //                        finalSplittedArray[finalSplittedArray.Length - 2] = stateZip[0];
        //                        finalSplittedArray[finalSplittedArray.Length - 1] = stateZip[1];
        //                    }
        //                    else if (lastWord.Contains("-"))
        //                    {
        //                        string[] stateZip = lastWord.Split('-');
        //                        finalSplittedArray[finalSplittedArray.Length - 2] = stateZip[0];
        //                        finalSplittedArray[finalSplittedArray.Length - 1] = stateZip[1];
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                finalSplittedArray = newSplittedArray;
        //            }

        //            // Step 2 - Starting from last word till first word, parse and find out sections as per the business rules
        //            // Last word => possibly the Zip Code -> If not then go to previous word until Zip Code is found
        //            // Set Zipcode as base word for parsing
        //            // Next check for State Code / State
        //            // Next check for City Name
        //            // Next all words in Address line 1 / 2
        //            // Next direction letter -> E, W, N, S
        //            // Last Street Number / House Number / Building Number
        //            bool isZipCodeFound = false;
        //            bool isStateFound = false;
        //            bool isCityFound = false;
        //            bool isAddressLine1Found = false;
        //            bool isDirectionFound = false;
        //            bool isStreetNumberFound = false;
        //            StringBuilder addressLine = new StringBuilder();

        //            for (int index = finalSplittedArray.Length - 1; index >= 0; index--)
        //            {
        //                string currentWord = finalSplittedArray[index];

        //                if (!string.IsNullOrEmpty(currentWord))
        //                {
        //                    // Remove special characters from start and end of word and then compare
        //                    // REM Remove all intial special character from the string
        //                    string strWord = Regex.Replace(currentWord, "^[^a-zA-Z0-9-]+", "");
        //                    // REM Remove all ending special character from the string
        //                    strWord = Regex.Replace(strWord, "[^a-zA-Z0-9-]*$", "");

        //                    // Search for ZipCode
        //                    if (isZipCodeFound == false)
        //                    {
        //                        if (strWord.Length >= 3) // Zip code could be of 3, 4, 5 OR 9 digits length
        //                        {
        //                            Match match = Regex.Match(strWord, "^[0-9-]+$");
        //                            if (match.Success == true)
        //                            {
        //                                // Zip Code is found -  set the value and move to next word
        //                                // Verify the value from database => if valid Zip code is found

        //                                // If the Zip Code length is more than 5, consider only first 5 digits for database comparison
        //                                if (strWord.Length > 5)
        //                                {
        //                                    strWord = strWord.Substring(0, 5);
        //                                }

        //                                isZipCodeFound = ParseUSAddress_SearchZipCodeInUSADatabase(strWord, SQLConnectionObject);
        //                                if (isZipCodeFound == true)
        //                                {
        //                                    objAddressSections.ZipCode = strWord;
        //                                    continue;
        //                                }
        //                            }
        //                        }
        //                    }
        //                    // Search for State Code / State Name
        //                    if (isZipCodeFound == true && isStateFound == false)
        //                    {
        //                        string[] usaStateCodes = {
        //                "AA", "AE", "AP", "AK", "AL", "AR", "AZ", "CA", "CO", "CT", "DC", "DE", "FL", "GA", "GU", "HI", "IA", "ID", "IL", "IN",
        //                "KS", "KY", "LA", "MA", "MD", "ME", "MI", "MN", "MO", "MS", "MT", "NC", "ND", "NE", "NH", "NJ", "NM", "NV", "NY", "OH",
        //                "OK", "OR", "PA", "PR", "RI", "SC", "SD", "TN", "TX", "UT", "VA", "VI", "VT", "WA", "WI", "WV", "WY"};

        //                        // USA State Names -  Full Names in upper case
        //                        string[] usaStateCodesFullNames = {
        //                "ARMED FORCES AMERICA", "ARMED FORCES", "ALASKA", "ALABAMA", "ARMED FORCES PACIFIC", "ARKANSAS",
        //                    "AMERICAN SAMOA", "ARIZONA", "CALIFORNIA", "COLORADO", "CONNECTICUT", "WASHINGTON DC", "DELAWARE", "FLORIDA", "GEORGIA", "GUAM", "HAWAII",
        //                    "IOWA", "IDAHO", "ILLINOIS", "INDIANA", "KANSAS", "KENTUCKY", "LOUISIANA", "MASSACHUSETTS", "MARYLAND", "MAINE", "MICHIGAN", "MINNESOTA",
        //                    "MISSOURI", "MISSISSIPPI", "MONTANA", "NORTH CAROLINA", "NORTH DAKOTA", "NEBRASKA", "NEW HAMPSHIRE", "NEW JERSEY", "NEW MEXICO", "NEVADA",
        //                    "NEW YORK", "OHIO", "OKLAHOMA", "OREGON", "PENNSYLVANIA", "PUERTO RICO", "RHODE ISLAND", "SOUTH CAROLINA", "SOUTH DAKOTA", "TENNESSEE",
        //                    "TEXAS", "UTAH", "VIRGINIA", "VIRGIN ISLANDS", "VERMONT", "WASHINGTON", "WISCONSIN", "WEST VIRGINIA", "WYOMING"};

        //                        if (Array.IndexOf(usaStateCodes, strWord.ToUpper()) >= 0 || Array.IndexOf(usaStateCodesFullNames, strWord.ToUpper()) >= 0)
        //                        {
        //                            // State Code / State is found -  set the value and move to next word
        //                            // Verify the value from database
        //                            // if state code length is more than 2, get state code value from database using state name and then search in DB
        //                            string stateCode = "";
        //                            if (strWord.Length > 2)
        //                            {
        //                                stateCode = ParseUSAddress_SearchUSAStateCodeUsingStateName(strWord, SQLConnectionObject);
        //                            }
        //                            else
        //                            {
        //                                stateCode = strWord;
        //                            }
        //                            isStateFound = ParseUSAddress_SearchStateCodeInUSADatabase(stateCode, SQLConnectionObject);
        //                            if (isStateFound == true)
        //                            {
        //                                objAddressSections.State = stateCode;
        //                                continue;
        //                            }
        //                        }
        //                    }
        //                    // Search for City
        //                    if (isZipCodeFound == true && isStateFound == true && isCityFound == false)
        //                    {
        //                        string zipCode = objAddressSections.ZipCode;
        //                        if (zipCode.Length > 5)
        //                        {
        //                            zipCode = zipCode.Substring(0, 5);
        //                        }
        //                        string cityFoundValue = ParseUSAddress_SearchCityInUSADatabaseUsingZipAndState(zipCode, objAddressSections.State, strWord, SQLConnectionObject);
        //                        if (!string.IsNullOrEmpty(cityFoundValue))
        //                        {
        //                            // Compare and extract the words from address line words
        //                            if (strWord.ToUpper() == cityFoundValue.ToUpper())
        //                            {
        //                                objAddressSections.City = strWord;
        //                                isCityFound = true;
        //                                continue;
        //                            }
        //                            else
        //                            {
        //                                string combinedWord = strWord;
        //                                if (finalSplittedArray[index - 1] != null)
        //                                {
        //                                    string previousWord = finalSplittedArray[index - 1];
        //                                    combinedWord = string.Concat(strWord, " ", previousWord);
        //                                    combinedWord = string.Join(" ", combinedWord.Split(' ').Reverse().ToArray());
        //                                    if (combinedWord.ToUpper() == cityFoundValue.ToUpper())
        //                                    {
        //                                        objAddressSections.City = combinedWord;
        //                                        isCityFound = true;
        //                                        index = index - 1;
        //                                        continue;
        //                                    }
        //                                    else
        //                                    {
        //                                        if (finalSplittedArray[index - 2] != null)
        //                                        {
        //                                            string previousWord1 = finalSplittedArray[index - 2];
        //                                            combinedWord = string.Concat(previousWord1, " ", combinedWord);
        //                                            //combinedWord = string.Join(" ", combinedWord.Split(' ').Reverse().ToArray());
        //                                            if (combinedWord.ToUpper() == cityFoundValue.ToUpper())
        //                                            {
        //                                                objAddressSections.City = combinedWord;
        //                                                isCityFound = true;
        //                                                index = index - 2;
        //                                                continue;
        //                                            }
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                        }
        //                    }
        //                    // Search for Address Line1
        //                    if (isZipCodeFound == true && isStateFound == true && isCityFound == true && isAddressLine1Found == false)
        //                    {
        //                        // check all the words until single letter direction OR numeric word is found which will be direction or street number
        //                        Match match = Regex.Match(strWord, "^[0-9-]+$");
        //                        if (strWord != "E" && strWord != "W" && strWord != "N" && strWord != "S" && match.Success == false)
        //                        {
        //                            if (!string.IsNullOrEmpty(Convert.ToString(addressLine)))
        //                            {
        //                                addressLine.Append(" ");
        //                            }
        //                            addressLine.Append(strWord);
        //                        }
        //                        else
        //                        {
        //                            // Address Line1 is found -  set the value and move to next word
        //                            objAddressSections.AddressLine1 = string.Join(" ", Convert.ToString(addressLine).Split(' ').Reverse().ToArray());
        //                            isAddressLine1Found = true;
        //                        }
        //                    }
        //                    // Search for Direction
        //                    if (isZipCodeFound == true && isStateFound == true && isCityFound == true && isAddressLine1Found == true && isDirectionFound == false)
        //                    {
        //                        Match match = Regex.Match(strWord, "[NEWS]+");
        //                        if (strWord.Length == 1 && match.Success == true)
        //                        {
        //                            // Direction is found -  set the value and move to next word
        //                            objAddressSections.Direction = strWord;
        //                            isDirectionFound = true;
        //                        }
        //                        else // Direction may be or may not be present in address, if not present and we found street number, skip the direction
        //                        {
        //                            Match match1 = Regex.Match(strWord, "^[0-9]+$");
        //                            if (match1.Success == true)
        //                            {
        //                                isDirectionFound = true;
        //                            }
        //                        }
        //                    }
        //                    // Search for Street Number
        //                    if (isZipCodeFound == true && isStateFound == true && isCityFound == true && isAddressLine1Found == true && isDirectionFound == true && isStreetNumberFound == false)
        //                    {
        //                        Match match = Regex.Match(strWord, "^[0-9]+$");
        //                        if (match.Success == true)
        //                        {
        //                            // Street Number is found -  set the value and break the loop
        //                            objAddressSections.StreetNumber = strWord;
        //                            isStreetNumberFound = true;
        //                            break;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //    return objAddressSections;
        //}

        //private static bool ParseUSAddress_SearchZipCodeInUSADatabase(string ZipCodeValue, SqlConnection SQLConnectionObject)
        //{
        //    bool isZipCodeFound = false;
        //    try
        //    {
        //        string cmdString = "select * from USCityStateZip where zipcode = '" + ZipCodeValue + "'";
        //        using (SqlCommand comm = new SqlCommand())
        //        {
        //            comm.Connection = SQLConnectionObject;
        //            comm.CommandText = cmdString;
        //            try
        //            {
        //                SQLConnectionObject.Open();
        //                SqlDataAdapter da = new SqlDataAdapter(comm);
        //                DataTable dataTable = new DataTable();
        //                da.Fill(dataTable);
        //                if (dataTable != null && dataTable.Rows != null && dataTable.Rows.Count > 0)
        //                {
        //                    isZipCodeFound = true;
        //                }
        //                SQLConnectionObject.Close();
        //                da.Dispose();
        //            }
        //            catch
        //            {
        //                isZipCodeFound = false;
        //            }
        //        }
        //    }
        //    catch
        //    {
        //        isZipCodeFound = false;
        //    }
        //    return isZipCodeFound;
        //}

        //private static string ParseUSAddress_SearchUSAStateCodeUsingStateName(string StateName, SqlConnection SQLConnectionObject)
        //{
        //    string stateCodeFoundValue = "";
        //    try
        //    {
        //        string cmdString = "select * from USAStateCodeName where StateName = '" + StateName + "'";
        //        using (SqlCommand comm = new SqlCommand())
        //        {
        //            comm.Connection = SQLConnectionObject;
        //            comm.CommandText = cmdString;
        //            try
        //            {
        //                SQLConnectionObject.Open();
        //                SqlDataAdapter da = new SqlDataAdapter(comm);
        //                DataTable dataTable = new DataTable();
        //                da.Fill(dataTable);
        //                if (dataTable != null && dataTable.Rows != null && dataTable.Rows.Count > 0)
        //                {
        //                    stateCodeFoundValue = Convert.ToString(dataTable.Rows[0]["StateCode"]);
        //                }
        //                SQLConnectionObject.Close();
        //                da.Dispose();
        //            }
        //            catch
        //            {
        //                stateCodeFoundValue = "";
        //            }
        //        }
        //    }
        //    catch
        //    {
        //        stateCodeFoundValue = "";
        //    }
        //    return stateCodeFoundValue;
        //}

        //private static bool ParseUSAddress_SearchStateCodeInUSADatabase(string StateCodeValue, SqlConnection SQLConnectionObject)
        //{
        //    bool isStateCodeFound = false;
        //    try
        //    {
        //        string cmdString = "select * from USCityStateZip where state = '" + StateCodeValue + "'";

        //        using (SqlCommand comm = new SqlCommand())
        //        {
        //            comm.Connection = SQLConnectionObject;
        //            comm.CommandText = cmdString;
        //            try
        //            {
        //                SQLConnectionObject.Open();
        //                SqlDataAdapter da = new SqlDataAdapter(comm);
        //                DataTable dataTable = new DataTable();
        //                da.Fill(dataTable);
        //                if (dataTable != null && dataTable.Rows != null && dataTable.Rows.Count > 0)
        //                {
        //                    isStateCodeFound = true;
        //                }
        //                SQLConnectionObject.Close();
        //                da.Dispose();
        //            }
        //            catch
        //            {
        //                isStateCodeFound = false;
        //            }
        //        }
        //    }
        //    catch
        //    {
        //        isStateCodeFound = false;
        //    }
        //    return isStateCodeFound;
        //}

        //private static string ParseUSAddress_SearchCityInUSADatabaseUsingZipAndState(string ZipCode, string StateCode, string CityFromAddress, SqlConnection SQLConnectionObject)
        //{
        //    string cityFoundValue = "";
        //    try
        //    {
        //        string cmdString = "select * from USCityStateZip where zipcode = '" + ZipCode + "' AND state = '" + StateCode + "'";
        //        using (SqlCommand comm = new SqlCommand())
        //        {
        //            comm.Connection = SQLConnectionObject;
        //            comm.CommandText = cmdString;
        //            try
        //            {
        //                SQLConnectionObject.Open();
        //                SqlDataAdapter da = new SqlDataAdapter(comm);
        //                DataTable dataTable = new DataTable();
        //                da.Fill(dataTable);
        //                if (dataTable != null && dataTable.Rows != null && dataTable.Rows.Count > 0)
        //                {
        //                    if (dataTable.Rows.Count > 1)
        //                    {
        //                        foreach (DataRow dr in dataTable.Rows)
        //                        {
        //                            if (Convert.ToString(dr["city"]).ToUpper().Contains(CityFromAddress.ToUpper()))
        //                            {
        //                                cityFoundValue = Convert.ToString(dr["city"]);
        //                                break;
        //                            }
        //                        }
        //                    }
        //                    else
        //                    {
        //                        cityFoundValue = Convert.ToString(dataTable.Rows[0]["city"]);
        //                    }
        //                }
        //                SQLConnectionObject.Close();
        //                da.Dispose();
        //            }
        //            catch
        //            {
        //                cityFoundValue = "";
        //            }
        //        }
        //    }
        //    catch
        //    {
        //        cityFoundValue = "";
        //    }
        //    return cityFoundValue;
        //}
        #endregion

        public static AddressSection ParseUSAddress(string AddressLineTextValue, DataTable USCityStateZipTable, DataTable USAStateCodeNameTable)
        {
            AddressSection objAddressSections = new AddressSection();
            try
            {
                string AddressLineText = AddressLineTextValue;

                // Check if address line contains "PO Box" keywords => "PO Box", "P.O. Box", P.O.Box", remove all PO Box keywords and the immediate number
                // Commented PO BOX removal code as discussed on 28th Jan 2021 => will keep PO Box and number for address comparison
                //bool isPOBOXPresent = false;
                bool isSuitePresent = false;
                //if (AddressLineText.ToUpper().Contains("PO BOX") || AddressLineText.ToUpper().Contains("P.O. BOX") ||
                //    AddressLineText.ToUpper().Contains("P.O.BOX") || AddressLineText.ToUpper().Contains("P O BOX"))
                //{
                //    isPOBOXPresent = true;
                //}

                // Check if address line contains "Suite" keywords => "Suite", "STE" remove all Suite keywords and the immediate number
                if (AddressLineText.ToUpper().Contains("SUITE") || AddressLineText.ToUpper().Contains("STE"))
                {
                    isSuitePresent = true;
                }

                // Step 1 - Split the address line text into separate words using "single space"
                string[] splittedArray = AddressLineText.Split(' ');
                string[] newSplittedArray = new string[splittedArray.Length + 1];

                // Remove PO BOX, Suite keywords and immediate next word from address
                //if (isPOBOXPresent || isSuitePresent)
                if (isSuitePresent)
                {
                    //if (isPOBOXPresent)
                    //{
                    //    int j = 0;
                    //    for (int index = 0; index < splittedArray.Length; index++)
                    //    {
                    //        string word = splittedArray[index];
                    //        if (word.ToUpper() == "PO" || word.ToUpper() == "P.O.")
                    //        {
                    //            index = index + 2;
                    //        }
                    //        else if (word.ToUpper() == "P.O.BOX")
                    //        {
                    //            index = index + 1;
                    //        }
                    //        else if (word.ToUpper() == "P")
                    //        {
                    //            if (index + 1 < splittedArray.Length)
                    //            {
                    //                if (splittedArray[index + 1].ToUpper() == "O")
                    //                {
                    //                    index = index + 3;
                    //                }
                    //            }
                    //        }
                    //        else
                    //        {
                    //            newSplittedArray[j] = word;
                    //            j++;
                    //        }
                    //    }
                    //}
                    //else if (isSuitePresent)
                    if (isSuitePresent)
                    {
                        int j = 0;
                        for (int index = 0; index < splittedArray.Length; index++)
                        {
                            string word = splittedArray[index];
                            if (word.ToUpper() == "SUITE" || word.ToUpper() == "STE")
                            {
                                index = index + 1;
                            }
                            else
                            {
                                newSplittedArray[j] = word;
                                j++;
                            }
                        }
                    }
                }
                else
                {
                    newSplittedArray = splittedArray;
                }

                newSplittedArray = newSplittedArray.Where(item => item != null).ToArray();

                if (newSplittedArray.Length > 0)
                {
                    // Check if the last word of the array is only numeric zip code or contains combined state code with zip code
                    // if yes, split the state code and zip code into 2 separate words - state code and zip code
                    string[] finalSplittedArray = new string[newSplittedArray.Length + 1];

                    string lastWord = newSplittedArray[newSplittedArray.Length - 1];
                    Match matchNew = Regex.Match(lastWord, "^[0-9-]+$");
                    if (matchNew.Success == false)
                    {
                        int digitCount = lastWord.Count(item => Char.IsDigit(item));
                        if (digitCount == 5 || digitCount == 9 || digitCount == 3 || digitCount == 4) // Zip code could be of 3, 4, 5 or 9 digits length
                        {
                            for (int rowIndex = 0; rowIndex < newSplittedArray.Length; rowIndex++)
                            {
                                finalSplittedArray[rowIndex] = newSplittedArray[rowIndex];
                            }

                            if (lastWord.Contains(","))
                            {
                                string[] stateZip = lastWord.Split(',');
                                finalSplittedArray[finalSplittedArray.Length - 2] = stateZip[0];
                                finalSplittedArray[finalSplittedArray.Length - 1] = stateZip[1];
                            }
                            else if (lastWord.Contains("-"))
                            {
                                string[] stateZip = lastWord.Split('-');
                                finalSplittedArray[finalSplittedArray.Length - 2] = stateZip[0];
                                finalSplittedArray[finalSplittedArray.Length - 1] = stateZip[1];
                            }
                        }
                    }
                    else
                    {
                        finalSplittedArray = newSplittedArray;
                    }

                    // Step 2 - Starting from last word till first word, parse and find out sections as per the business rules
                    // Last word => possibly the Zip Code -> If not then go to previous word until Zip Code is found
                    // Set Zipcode as base word for parsing
                    // Next check for State Code / State
                    // Next check for City Name
                    // Next all words in Address line 1 / 2
                    // Next direction letter -> E, W, N, S
                    // Last Street Number / House Number / Building Number
                    bool isZipCodeFound = false;
                    bool isStateFound = false;
                    bool isCityFound = false;
                    bool isAddressLine1Found = false;
                    bool isDirectionFound = false;
                    bool isStreetNumberFound = false;
                    bool isNoCityPresent = false;
                    StringBuilder addressLine = new StringBuilder();
                    int loopIndex = -1;
                    for (int index = finalSplittedArray.Length - 1; index >= 0; index--)
                    {
                        string currentWord = finalSplittedArray[index];

                        if (!string.IsNullOrEmpty(currentWord))
                        {
                            // Remove special characters from start and end of word and then compare
                            // REM Remove all intial special character from the string
                            string strWord = Regex.Replace(currentWord, "^[^a-zA-Z0-9-]+", "");
                            // REM Remove all ending special character from the string
                            strWord = Regex.Replace(strWord, "[^a-zA-Z0-9-]*$", "");

                            // Search for ZipCode
                            if (isZipCodeFound == false)
                            {
                                if (strWord.Length >= 3) // Zip code could be of 3, 4, 5 OR 9 digits length
                                {
                                    Match match = Regex.Match(strWord, "^[0-9-]+$");
                                    if (match.Success == true)
                                    {
                                        // Zip Code is found -  set the value and move to next word
                                        // Verify the value from database => if valid Zip code is found

                                        // If the Zip Code length is more than 5, consider only first 5 digits for database comparison
                                        if (strWord.Length > 5)
                                        {
                                            strWord = strWord.Substring(0, 5);
                                        }

                                        isZipCodeFound = ParseUSAddress_SearchZipCodeInUSADatabase(strWord, USCityStateZipTable);
                                        if (isZipCodeFound == true)
                                        {
                                            objAddressSections.ZipCode = strWord;
                                            continue;
                                        }
                                    }
                                }
                            }
                            // Search for State Code / State Name
                            if (isZipCodeFound == true && isStateFound == false)
                            {
                                string[] usaStateCodes = {
                        "AA", "AE", "AP", "AK", "AL", "AR", "AZ", "CA", "CO", "CT", "DC", "DE", "FL", "GA", "GU", "HI", "IA", "ID", "IL", "IN",
                        "KS", "KY", "LA", "MA", "MD", "ME", "MI", "MN", "MO", "MS", "MT", "NC", "ND", "NE", "NH", "NJ", "NM", "NV", "NY", "OH",
                        "OK", "OR", "PA", "PR", "RI", "SC", "SD", "TN", "TX", "UT", "VA", "VI", "VT", "WA", "WI", "WV", "WY"};

                                // USA State Names -  Full Names in upper case
                                string[] usaStateCodesFullNames = {
                        "ARMED FORCES AMERICA", "ARMED FORCES", "ALASKA", "ALABAMA", "ARMED FORCES PACIFIC", "ARKANSAS",
                            "AMERICAN SAMOA", "ARIZONA", "CALIFORNIA", "COLORADO", "CONNECTICUT", "WASHINGTON DC", "DELAWARE", "FLORIDA", "GEORGIA", "GUAM", "HAWAII",
                            "IOWA", "IDAHO", "ILLINOIS", "INDIANA", "KANSAS", "KENTUCKY", "LOUISIANA", "MASSACHUSETTS", "MARYLAND", "MAINE", "MICHIGAN", "MINNESOTA",
                            "MISSOURI", "MISSISSIPPI", "MONTANA", "NORTH CAROLINA", "NORTH DAKOTA", "NEBRASKA", "NEW HAMPSHIRE", "NEW JERSEY", "NEW MEXICO", "NEVADA",
                            "NEW YORK", "OHIO", "OKLAHOMA", "OREGON", "PENNSYLVANIA", "PUERTO RICO", "RHODE ISLAND", "SOUTH CAROLINA", "SOUTH DAKOTA", "TENNESSEE",
                            "TEXAS", "UTAH", "VIRGINIA", "VIRGIN ISLANDS", "VERMONT", "WASHINGTON", "WISCONSIN", "WEST VIRGINIA", "WYOMING"};

                                if (Array.IndexOf(usaStateCodes, strWord.ToUpper()) >= 0 || Array.IndexOf(usaStateCodesFullNames, strWord.ToUpper()) >= 0)
                                {
                                    // State Code / State is found -  set the value and move to next word
                                    // Verify the value from database
                                    // if state code length is more than 2, get state code value from database using state name and then search in DB
                                    string stateCode = "";
                                    if (strWord.Length > 2)
                                    {
                                        stateCode = ParseUSAddress_SearchUSAStateCodeUsingStateName(strWord, USAStateCodeNameTable);
                                    }
                                    else
                                    {
                                        stateCode = strWord;
                                    }
                                    isStateFound = ParseUSAddress_SearchStateCodeInUSADatabase(stateCode, USCityStateZipTable);
                                    if (isStateFound == true)
                                    {
                                        objAddressSections.State = stateCode;
                                        continue;
                                    }
                                }
                            }
                            // Search for City
                            if (isZipCodeFound == true && isStateFound == true && isCityFound == false)
                            {
                                // If city is not found, this will be used to find other remaining sections
                                loopIndex = index;

                                string zipCode = objAddressSections.ZipCode;
                                if (zipCode.Length > 5)
                                {
                                    zipCode = zipCode.Substring(0, 5);
                                }
                                string cityFoundValue = ParseUSAddress_SearchCityInUSADatabaseUsingZipAndState(zipCode, objAddressSections.State, strWord, USCityStateZipTable);
                                if (!string.IsNullOrEmpty(cityFoundValue))
                                {
                                    // Compare and extract the words from address line words
                                    if (strWord.ToUpper() == cityFoundValue.ToUpper())
                                    {
                                        objAddressSections.City = strWord;
                                        isCityFound = true;
                                        loopIndex = index;
                                        break;
                                        //continue;
                                    }
                                    else
                                    {
                                        string combinedWord = strWord;
                                        if (index > 0) // Added on 27th Jan 2021 -> to handle the condition, where no more words are present to left of current word
                                        {
                                            if (finalSplittedArray[index - 1] != null)
                                            {
                                                string previousWord = finalSplittedArray[index - 1];
                                                combinedWord = string.Concat(strWord, " ", previousWord);
                                                combinedWord = string.Join(" ", combinedWord.Split(' ').Reverse().ToArray());
                                                if (combinedWord.ToUpper() == cityFoundValue.ToUpper())
                                                {
                                                    objAddressSections.City = combinedWord;
                                                    isCityFound = true;
                                                    index = index - 1;
                                                    loopIndex = index;
                                                    break;
                                                    //continue;
                                                }
                                                else
                                                {
                                                    if (finalSplittedArray[index - 2] != null)
                                                    {
                                                        string previousWord1 = finalSplittedArray[index - 2];
                                                        combinedWord = string.Concat(previousWord1, " ", combinedWord);
                                                        //combinedWord = string.Join(" ", combinedWord.Split(' ').Reverse().ToArray());
                                                        if (combinedWord.ToUpper() == cityFoundValue.ToUpper())
                                                        {
                                                            objAddressSections.City = combinedWord;
                                                            isCityFound = true;
                                                            index = index - 2;
                                                            loopIndex = index;
                                                            break;
                                                            //continue;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            // if city is no found
                            if (isZipCodeFound == true && isStateFound == true && isCityFound == false)
                            {
                                isNoCityPresent = true;
                                break;
                            }

                            //// If 3 sections are found, break the loop and then find house # first, direction and then 
                            //if (isZipCodeFound == true && isStateFound == true && isCityFound == true)
                            //{
                            //    loopIndex = index;
                            //    break;
                            //}
                            //// Search for Address Line1
                            //if (isZipCodeFound == true && isStateFound == true && isCityFound == true && isAddressLine1Found == false)
                            //{
                            //    // check all the words until single letter direction OR numeric word is found which will be direction or street number
                            //    Match match = Regex.Match(strWord, "^[0-9-]+$");
                            //    if (strWord != "E" && strWord != "W" && strWord != "N" && strWord != "S" && match.Success == false)
                            //    {
                            //        if (!string.IsNullOrEmpty(Convert.ToString(addressLine)))
                            //        {
                            //            addressLine.Append(" ");
                            //        }
                            //        addressLine.Append(strWord);
                            //    }
                            //    else
                            //    {
                            //        // Address Line1 is found -  set the value and move to next word
                            //        objAddressSections.AddressLine1 = string.Join(" ", Convert.ToString(addressLine).Split(' ').Reverse().ToArray());
                            //        isAddressLine1Found = true;
                            //    }
                            //}
                            //// Search for Direction
                            //if (isZipCodeFound == true && isStateFound == true && isCityFound == true && isAddressLine1Found == true && isDirectionFound == false)
                            //{
                            //    Match match = Regex.Match(strWord, "[NEWS]+");
                            //    if (strWord.Length == 1 && match.Success == true)
                            //    {
                            //        // Direction is found -  set the value and move to next word
                            //        objAddressSections.Direction = strWord;
                            //        isDirectionFound = true;
                            //    }
                            //    else // Direction may be or may not be present in address, if not present and we found street number, skip the direction
                            //    {
                            //        Match match1 = Regex.Match(strWord, "^[0-9]+$");
                            //        if (match1.Success == true)
                            //        {
                            //            isDirectionFound = true;
                            //        }
                            //    }
                            //}
                            //// Search for Street Number
                            //if (isZipCodeFound == true && isStateFound == true && isCityFound == true && isAddressLine1Found == true && isDirectionFound == true && isStreetNumberFound == false)
                            //{
                            //    Match match = Regex.Match(strWord, "^[0-9]+$");
                            //    if (match.Success == true)
                            //    {
                            //        // Street Number is found -  set the value and break the loop
                            //        objAddressSections.StreetNumber = strWord;
                            //        isStreetNumberFound = true;
                            //        break;
                            //    }
                            //}
                        }
                    }

                    // Find House #, Direction and Address Line 1
                    // Added on 11th Feb 2021 = > check if street number is found from P O BOX address line and if BOX is present just before street number, set Street number to 0
                    int streetNumberIndex = -1;
                    if (loopIndex >= 0)
                    {
                        for (int index = 0; index < loopIndex; index++)
                        {
                            string currentWord = finalSplittedArray[index];
                            if (!string.IsNullOrEmpty(currentWord))
                            {
                                // Remove special characters from start and end of word and then compare
                                // REM Remove all intial special character from the string
                                string strWord = Regex.Replace(currentWord, "^[^a-zA-Z0-9-]+", "");
                                // REM Remove all ending special character from the string
                                strWord = Regex.Replace(strWord, "[^a-zA-Z0-9-]*$", "");

                                // Search for Street Number
                                if (isZipCodeFound == true && isStateFound == true && (isCityFound == true || isNoCityPresent == true) && isStreetNumberFound == false)
                                {
                                    Match match = Regex.Match(strWord, "^[0-9]+$");
                                    if (match.Success == true)
                                    {
                                        // Street Number is found -  set the value and break the loop
                                        objAddressSections.StreetNumber = strWord;
                                        isStreetNumberFound = true;
                                        streetNumberIndex = index;  // Added on 11th Feb 2021
                                        continue;
                                    }
                                }

                                // Search for Direction
                                if (isZipCodeFound == true && isStateFound == true && (isCityFound == true || isNoCityPresent == true) && isStreetNumberFound == true && isDirectionFound == false)
                                {
                                    Match match = Regex.Match(strWord, "[NEWS]+");
                                    if (strWord.Length == 1 && match.Success == true)
                                    {
                                        // Direction is found -  set the value and move to next word
                                        objAddressSections.Direction = strWord;
                                        isDirectionFound = true;
                                        continue;
                                    }
                                    else if (strWord.Length == 2 && (strWord == "NE" || strWord == "SE" || strWord == "NW" || strWord == "SW"))
                                    {
                                        // Direction is found -  set the value and move to next word
                                        objAddressSections.Direction = strWord;
                                        isDirectionFound = true;
                                        continue;
                                    }
                                    else if ((strWord.Length == 4 || strWord.Length == 5) && (strWord == "EAST" || strWord == "WEST" || strWord == "NORTH" || strWord == "SOUTH"))
                                    {
                                        // Direction is found -  set the value and move to next word
                                        objAddressSections.Direction = strWord;
                                        isDirectionFound = true;
                                        continue;
                                    }
                                    else // Direction may be or may not be present in address, if not present , skip the direction
                                    {
                                        addressLine.Append(strWord);
                                        isDirectionFound = true;
                                        continue;
                                        //Match match1 = Regex.Match(strWord, "^[0-9]+$");
                                        //if (match1.Success == true)
                                        //{
                                        //    isDirectionFound = true;
                                        //}
                                    }
                                }

                                // Search for Address Line1
                                if (isZipCodeFound == true && isStateFound == true && (isCityFound == true || isNoCityPresent == true) && isStreetNumberFound == true &&
                                    isDirectionFound == true && isAddressLine1Found == false)
                                {
                                    //// check all the words until single letter direction OR numeric word is found which will be direction or street number
                                    //Match match = Regex.Match(strWord, "^[0-9-]+$");
                                    //if (strWord != "E" && strWord != "W" && strWord != "N" && strWord != "S" && match.Success == false)
                                    //{
                                    //    if (!string.IsNullOrEmpty(Convert.ToString(addressLine)))
                                    //    {
                                    //        addressLine.Append(" ");
                                    //    }
                                    //    addressLine.Append(strWord);
                                    //}
                                    //else
                                    //{
                                    //    // Address Line1 is found -  set the value and move to next word
                                    //    objAddressSections.AddressLine1 = string.Join(" ", Convert.ToString(addressLine).Split(' ').Reverse().ToArray());
                                    //    isAddressLine1Found = true;
                                    //}
                                    if (!string.IsNullOrEmpty(Convert.ToString(addressLine)))
                                    {
                                        addressLine.Append(" ");
                                    }
                                    addressLine.Append(strWord);
                                }

                            }
                        }

                        // If address line is empty and stree #, direction and address line is not found, consider all remaining words as address line
                        if (string.IsNullOrEmpty(Convert.ToString(addressLine)))
                        {
                            for (int index = 0; index < loopIndex; index++)
                            {
                                if (!string.IsNullOrEmpty(Convert.ToString(addressLine)))
                                {
                                    addressLine.Append(" ");
                                }
                                addressLine.Append(finalSplittedArray[index]);
                            }
                        }
                        objAddressSections.AddressLine1 = Convert.ToString(addressLine);
                        isAddressLine1Found = true;

                        // Added on 11th Feb 2021
                        // Index should be greater than 0, as we need to check previous word, and if previous word is BOX, set Street number to 0
                        // In DB address, for all PO BOX addresses, Street number is 0
                        if (isStreetNumberFound == true && streetNumberIndex > 0)
                        {
                            if (finalSplittedArray[streetNumberIndex - 1].ToUpper().Contains("BOX"))
                            {
                                objAddressSections.StreetNumber = "0";
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            // Set the flag IsAddressParsed to FALSE, if all the values are NULL, else set it to TRUE => Added on 25th Jan 2021 to handle ALL NULL conditions for values
            if (objAddressSections.StreetNumber == null && objAddressSections.Direction == null && objAddressSections.AddressLine1 == null &&
                objAddressSections.City == null && objAddressSections.State == null && objAddressSections.ZipCode == null)
            {
                objAddressSections.IsAddressParsed = false;
            }
            else
            {
                objAddressSections.IsAddressParsed = true;
            }
            return objAddressSections;
        }

        private static bool ParseUSAddress_SearchZipCodeInUSADatabase(string ZipCodeValue, DataTable USCityStateZipTable)
        {
            bool isZipCodeFound = false;
            try
            {
                ////string cmdString = "select * from USCityStateZip where zipcode = '" + ZipCodeValue + "'";
                //string cmdString = "select * from USCityStateZip";
                //OdbcConnectionObject.propConnection.Open();
                //DataTable dataTable = OdbcConnectionObject.ReturnDataTable(cmdString, "USCityStateZip");
                if (USCityStateZipTable != null && USCityStateZipTable.Rows != null && USCityStateZipTable.Rows.Count > 0)
                {
                    DataRow[] drRows = USCityStateZipTable.Select("[zipcode] = '" + ZipCodeValue + "'");
                    if (drRows != null && drRows.Length > 0)
                    {
                        isZipCodeFound = true;
                    }
                }
                //OdbcConnectionObject.propConnection.Close();
                //OdbcConnectionObject = null;
            }
            catch
            {
                isZipCodeFound = false;
            }
            return isZipCodeFound;
        }

        private static string ParseUSAddress_SearchUSAStateCodeUsingStateName(string StateName, DataTable USAStateCodeNameTable)
        {
            string stateCodeFoundValue = "";
            try
            {
                ////string cmdString = "select * from USAStateCodeName where StateName = '" + StateName + "'";
                //string cmdString = "select * from USAStateCodeName";
                //OdbcConnectionObject.propConnection.Open();
                //DataTable dataTable = OdbcConnectionObject.ReturnDataTable(cmdString, "USAStateCodeName");
                if (USAStateCodeNameTable != null && USAStateCodeNameTable.Rows != null && USAStateCodeNameTable.Rows.Count > 0)
                {
                    DataRow[] drRows = USAStateCodeNameTable.Select("[StateName] = '" + StateName + "'");
                    if (drRows != null && drRows.Length > 0)
                    {
                        stateCodeFoundValue = Convert.ToString(drRows[0]["StateCode"]);
                    }
                }
                //OdbcConnectionObject.propConnection.Close();
                //OdbcConnectionObject = null;
            }
            catch
            {
                stateCodeFoundValue = "";
            }
            return stateCodeFoundValue;
        }

        private static bool ParseUSAddress_SearchStateCodeInUSADatabase(string StateCodeValue, DataTable USCityStateZipTable)
        {
            bool isStateCodeFound = false;
            try
            {
                ////string cmdString = "select * from USCityStateZip where state = '" + StateCodeValue + "'";
                //string cmdString = "select * from USCityStateZip";
                //OdbcConnectionObject.propConnection.Open();
                //DataTable dataTable = OdbcConnectionObject.ReturnDataTable(cmdString, "USCityStateZip");
                if (USCityStateZipTable != null && USCityStateZipTable.Rows != null && USCityStateZipTable.Rows.Count > 0)
                {
                    DataRow[] drRows = USCityStateZipTable.Select("[state] = '" + StateCodeValue + "'");
                    if (drRows != null && drRows.Length > 0)
                    {
                        isStateCodeFound = true;
                    }
                }
                //OdbcConnectionObject.propConnection.Close();
                //OdbcConnectionObject = null;
            }
            catch
            {
                isStateCodeFound = false;
            }
            return isStateCodeFound;
        }

        private static string ParseUSAddress_SearchCityInUSADatabaseUsingZipAndState(string ZipCode, string StateCode, string CityFromAddress, DataTable USCityStateZipTable)
        {
            string cityFoundValue = "";
            try
            {
                ////string cmdString = "select * from USCityStateZip where zipcode = '" + ZipCode + "' AND state = '" + StateCode + "'";
                //string cmdString = "select * from USCityStateZip";
                //OdbcConnectionObject.propConnection.Open();
                //DataTable dataTable = OdbcConnectionObject.ReturnDataTable(cmdString, "USCityStateZip");
                if (USCityStateZipTable != null && USCityStateZipTable.Rows != null && USCityStateZipTable.Rows.Count > 0)
                {
                    DataRow[] drRows = USCityStateZipTable.Select("[zipcode] = '" + ZipCode + "' AND [state] = '" + StateCode + "'");
                    if (drRows != null && drRows.Length > 0)
                    {
                        if (drRows.Length > 1)
                        {
                            foreach (DataRow dr in drRows)
                            {
                                if (Convert.ToString(dr["city"]).ToUpper().Contains(CityFromAddress.ToUpper()))
                                {
                                    cityFoundValue = Convert.ToString(dr["city"]);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            cityFoundValue = Convert.ToString(drRows[0]["city"]);
                        }
                    }
                }
                //OdbcConnectionObject.propConnection.Close();
                //OdbcConnectionObject = null;
            }
            catch
            {
                cityFoundValue = "";
            }
            return cityFoundValue;
        }

        public static DataTable GetUSAZipCodeStateCityTable(clsOleDBDataAccess OdbcConnectionObject, string TableName)
        {
            string cmdString = string.Concat("select * from ", TableName);
            OdbcConnectionObject.propConnection.Open();
            DataTable dataTable = OdbcConnectionObject.ReturnDataTable(cmdString, TableName);
            OdbcConnectionObject.propConnection.Close();
            return dataTable;
        }
    }
    public class AddressSection
    {
        public string StreetNumber { get; set; }
        public string Direction { get; set; }
        public string AddressLine1 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        public bool IsAddressParsed { get; set; } // Added on 25th Jan 2021
    }
}
