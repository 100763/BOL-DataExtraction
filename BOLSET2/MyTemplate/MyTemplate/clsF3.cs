using iQDataProvider;
using iQPUBLIC;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text.RegularExpressions;
//using BOL_SET2;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.IO;
using System.Text;
using System.Collections;
//using System.Net;


namespace BOLSET2IQ
{
    public class clsF3
    {
        //BOL_SET2.clsF3 objBOLF3;
        public RetStructF3 F3(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strRoutineNo, string strArg)
        {
            // Added on 30-Dec-2021 => EDI Document
            ClsOCRValues.IsEDIDocument = "N";
            if (strArg.Contains('#'))
            {
                if (strArg.Split('#')[1] == "1")
                {
                    ClsOCRValues.IsEDIDocument = "Y";
                }
            }
            //RetStructF3 objRetStructF3 = new RetStructF3();
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            //if (Module1.imageName.ToLower() != Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath).ToLower())
            if (Module1.ActualimageName.ToLower() != Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath).ToLower())
            {
                Module1.TableRecordss = new DataTable();
                Module1.imageName = Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath).ToLower();
                if (Module1.dtUSCityStateZip.Rows.Count <= 0)
                {
                    Module1.dtUSCityStateZip = Module1.GetUSCityStateZipCode();
                }
            }
            switch (strRoutineNo)
            {
                #region USED 
                case "C800":// Get All Address from the image
                    if (!string.IsNullOrEmpty(iQPUBLIC.PublicComponents.PrimaryImagePath))
                    {
                        #region "Clear global varaibles"
                        // Shipper
                        ClsOCRValues.ClsOCRValuesShipperName = string.Empty;
                        ClsOCRValues.ClsOCRValuesShipperAddressLine1 = string.Empty;
                        ClsOCRValues.ClsOCRValuesShipperCity = string.Empty;
                        ClsOCRValues.ClsOCRValuesShipperState = string.Empty;
                        ClsOCRValues.ClsOCRValuesShipperZip = string.Empty;
                        ClsOCRValues.ClsOCRValuesShipperTelePhoneNo = string.Empty;
                        // Consignee
                        ClsOCRValues.ClsOCRValuesConsigneeName = string.Empty;
                        ClsOCRValues.ClsOCRValuesConsigneeAddressLine1 = string.Empty;
                        ClsOCRValues.ClsOCRValuesConsigneeCity = string.Empty;
                        ClsOCRValues.ClsOCRValuesConsigneeState = string.Empty;
                        ClsOCRValues.ClsOCRValuesConsigneeZip = string.Empty;
                        ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo = string.Empty;
                        // Bill to 3rd party
                        ClsOCRValues.ClsOCRValuesBillToName = string.Empty;
                        ClsOCRValues.ClsOCRValuesBillToAddressLine1 = string.Empty;
                        ClsOCRValues.ClsOCRValuesBillToCity = string.Empty;
                        ClsOCRValues.ClsOCRValuesBillToState = string.Empty;
                        ClsOCRValues.ClsOCRValuesBillToZip = string.Empty;
                        ClsOCRValues.ClsOCRValuesBillToTelePhoneNo = string.Empty;

                        // OCR Confidence level
                        // Name
                        ClsOCRValues.ShipperNameOCRWithHighConfidence = string.Empty;
                        ClsOCRValues.ConsigneeNameOCRWithHighConfidence = string.Empty;
                        ClsOCRValues.BillToNameOCRWithHighConfidence = string.Empty;
                        // Telephone
                        ClsOCRValues.ShipperTelephoneOCRWithHighConfidence = string.Empty;
                        ClsOCRValues.ConsigneeTelephoneOCRWithHighConfidence = string.Empty;
                        ClsOCRValues.BillToTelephoneOCRWithHighConfidence = string.Empty;

                        // Address
                        ClsOCRValues.ConsigneeAddressLine1OCRWithHighConfidence = string.Empty;
                        ClsOCRValues.ConsigneeCityOCRWithHighConfidence = string.Empty;
                        ClsOCRValues.ConsigneeStateOCRWithHighConfidence = string.Empty;
                        ClsOCRValues.ConsigneeZipOCRWithHighConfidence = string.Empty;
                        // Terms on OBL
                        ClsOCRValues.ClsOCRValuesTerms = string.Empty;
                        // Tag consignee only when if we found keyword or matched with Pick up DB
                        ClsOCRValues.ClsOCRCorrectConsigneeTagged = "F";  // Added on 6-July-2021 => Tag consignee only when if we found keyword or matched with Pick up DB
                        #endregion
                        if (ClsOCRValues.IsEDIDocument == "N") // Added on 30-Dec-2021 => EDI Document
                        {
                            Module1.objRetStructF3 = AllAddresses(ObjMetaData, intCurrPageNumber, strArg);
                        }
                        else
                        {
                            List<int> ConfLevel = new List<int>();
                            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
                            PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                            Module1.objRetStructF3.Words = PossibleWords;
                            ConfLevel.Add(90);
                            Module1.objRetStructF3.ManualConfirmation = "N";
                            Module1.objRetStructF3.Status = "S";
                            Module1.objRetStructF3.Flag = "Y";
                            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
                        }
                    }
                    break;
                case "C03": // Shipper Name => OCR
                    Module1.objRetStructF3 = ShipperNameOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C04": // Shipper address Line => OCR
                    Module1.objRetStructF3 = ShipperAddressLineOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C05": // Shipper City => OCR
                    Module1.objRetStructF3 = ShipperCityOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C06": // Shipper State => OCR
                    Module1.objRetStructF3 = ShipperStateOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C07": // Shipper Zip code => OCR
                    Module1.objRetStructF3 = ShipperZipOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C08": // Shipper TelNo => OCR
                    Module1.objRetStructF3 = ShipperTelNoOCR(ObjMetaData, intCurrPageNumber);
                    break;


                case "C033": // Consignee Name => OCR
                    Module1.objRetStructF3 = ConsigneeNameOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C044": // Consignee address Line => OCR
                    Module1.objRetStructF3 = ConsigneeAddressLineOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C055": // Consignee City => OCR
                    Module1.objRetStructF3 = ConsigneeCityOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C066": // Consignee State => OCR
                    Module1.objRetStructF3 = ConsigneeStateOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C077": // Consignee Zip code => OCR
                    Module1.objRetStructF3 = ConsigneeZipOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C088": // Consignee TelNo => OCR
                    Module1.objRetStructF3 = ConsingeeTelNoOCR(ObjMetaData, intCurrPageNumber);
                    break;

                // Bill To
                case "C0333": // Bill To Name => OCR
                    Module1.objRetStructF3 = BillToNameOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C0444": // Bill To address Line => OCR
                    Module1.objRetStructF3 = BillToAddressLineOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C0555": // Bill To City => OCR
                    Module1.objRetStructF3 = BillToCityOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C0666": // Bill To State => OCR
                    Module1.objRetStructF3 = BillToStateOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C0777": // Bill To Zip code => OCR
                    Module1.objRetStructF3 = BillToZipOCR(ObjMetaData, intCurrPageNumber);
                    break;
                case "C0888": // Bill To TelNo => OCR
                    Module1.objRetStructF3 = BillToTelNoOCR(ObjMetaData, intCurrPageNumber);
                    break;

                case "C013": // Terms
                    Module1.objRetStructF3 = F3_AutoLocate_Terms(ObjMetaData, intCurrPageNumber, strArg);
                    break;
                case "C014": // Terms OBL
                    Module1.objRetStructF3 = Terms_OBL(ObjMetaData, intCurrPageNumber, strArg);
                    break;
                    #endregion
            }

            return Module1.objRetStructF3;

        }

        #region USED
        // Shipper 
        private RetStructF3 ShipperNameOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesShipperName = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objShipperNameOCR = new clsCnCWord();
                    //objAdd1 = (clsCnCWord)Module1.TableRecordssFinal.Rows[0]["cncName"];
                    objShipperNameOCR = (clsCnCWord)dataRowsShipper[0]["cncName"];
                    PossibleWords.Add(objShipperNameOCR);
                    // check the character confidence 
                    var NineDigitcount = objShipperNameOCR.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                    var EightDigitcount = objShipperNameOCR.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
                    var ZeroDigitcount = objShipperNameOCR.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
                    if (objShipperNameOCR.Confidence >= 85 && (objShipperNameOCR.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
                    {
                        ClsOCRValues.ShipperNameOCRWithHighConfidence = "Y";
                    }
                    else
                    {
                        ClsOCRValues.ShipperNameOCRWithHighConfidence = "N";
                    }

                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else if (ClsOCRValues.IsEDIDocument == "Y") // Added on 30-Dec-2021 => EDI Document
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, clsF27.clsModule.clsPickupValuesFRP001ShipperName));
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }

            ClsOCRValues.ClsOCRValuesShipperName = PossibleWords[0].strWord;
            if (PossibleWords[0].strWord != null && PossibleWords[0].strWord.Trim().ToLower().Contains("in care of"))
            {
                clsF27.clsModule.clsPickupValuesFRP001RAT = "Y";
            }
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 ShipperAddressLineOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesShipperAddressLine1 = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objShipperAddressLineOCR = new clsCnCWord();
                    objShipperAddressLineOCR = (clsCnCWord)dataRowsShipper[0]["cncAddress"];
                    PossibleWords.Add(objShipperAddressLineOCR);
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else if (ClsOCRValues.IsEDIDocument == "Y") // Added on 30-Dec-2021 => EDI Document
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, clsF27.clsModule.clsPickupValuesFRP001ShipperAddress1));
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesShipperAddressLine1 = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 ShipperCityOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesShipperCity = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objShipperCityOCR = new clsCnCWord();
                    objShipperCityOCR = (clsCnCWord)dataRowsShipper[0]["cncCity"];
                    PossibleWords.Add(objShipperCityOCR);
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else if (ClsOCRValues.IsEDIDocument == "Y") // Added on 30-Dec-2021 => EDI Document
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, clsF27.clsModule.clsPickupValuesFRP001ShipperCity));
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesShipperCity = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 ShipperStateOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesShipperState = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objShipperStateOCR = new clsCnCWord();
                    objShipperStateOCR = (clsCnCWord)dataRowsShipper[0]["cncState"];
                    PossibleWords.Add(objShipperStateOCR);
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else if (ClsOCRValues.IsEDIDocument == "Y") // Added on 30-Dec-2021 => EDI Document
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, clsF27.clsModule.clsPickupValuesFRP001ShipperState));
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesShipperState = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 ShipperZipOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesShipperZip = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objShipperZipOCR = new clsCnCWord();
                    objShipperZipOCR = (clsCnCWord)dataRowsShipper[0]["cncZip"];
                    PossibleWords.Add(objShipperZipOCR);
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else if (ClsOCRValues.IsEDIDocument == "Y") // Added on 30-Dec-2021 => EDI Document
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, clsF27.clsModule.clsPickupValuesFRP001ShipperZip));
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesShipperZip = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 ShipperTelNoOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesShipperTelePhoneNo = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objShipperTelNoOCR = new clsCnCWord();
                    objShipperTelNoOCR = (clsCnCWord)dataRowsShipper[0]["cncTelPhone"];
                    if (objShipperTelNoOCR.strWord != null)
                    {
                        PossibleWords.Add(objShipperTelNoOCR);
                        // check the character confidence 
                        var NineDigitcount = objShipperTelNoOCR.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                        var EightDigitcount = objShipperTelNoOCR.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
                        var ZeroDigitcount = objShipperTelNoOCR.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
                        if (objShipperTelNoOCR.Confidence >= 85 && (objShipperTelNoOCR.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
                        {
                            ClsOCRValues.ShipperTelephoneOCRWithHighConfidence = "Y";
                        }
                        else
                        {
                            ClsOCRValues.ShipperTelephoneOCRWithHighConfidence = "N";
                        }
                    }
                    else
                    {
                        PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                    }
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesShipperTelePhoneNo = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        // Consignee
        private RetStructF3 ConsigneeNameOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesConsigneeName = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objConsigneeNameOCR = new clsCnCWord();
                    //objAdd1 = (clsCnCWord)Module1.TableRecordssFinal.Rows[0]["cncName"];
                    objConsigneeNameOCR = (clsCnCWord)dataRowsShipper[0]["cncName"];
                    PossibleWords.Add(objConsigneeNameOCR);
                    var NineDigitcount = objConsigneeNameOCR.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                    var EightDigitcount = objConsigneeNameOCR.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
                    var ZeroDigitcount = objConsigneeNameOCR.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
                    if (objConsigneeNameOCR.Confidence >= 85 && (objConsigneeNameOCR.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
                    {
                        ClsOCRValues.ConsigneeNameOCRWithHighConfidence = "Y";
                    }
                    else
                    {
                        ClsOCRValues.ConsigneeNameOCRWithHighConfidence = "N";
                    }
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else if (ClsOCRValues.IsEDIDocument == "Y") // Added on 30-Dec-2021 => EDI Document
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, clsF27.clsModule.clsPickupValuesFRP001ConsigneeName));
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesConsigneeName = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 ConsigneeAddressLineOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesConsigneeAddressLine1 = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objConsigneeAddressLineOCR = new clsCnCWord();
                    objConsigneeAddressLineOCR = (clsCnCWord)dataRowsShipper[0]["cncAddress"];
                    PossibleWords.Add(objConsigneeAddressLineOCR);

                    var NineDigitcount = objConsigneeAddressLineOCR.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                    var EightDigitcount = objConsigneeAddressLineOCR.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
                    var ZeroDigitcount = objConsigneeAddressLineOCR.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
                    if (objConsigneeAddressLineOCR.Confidence >= 85 && (objConsigneeAddressLineOCR.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
                    {
                        ClsOCRValues.ConsigneeAddressLine1OCRWithHighConfidence = "Y";
                    }
                    else
                    {
                        ClsOCRValues.ConsigneeAddressLine1OCRWithHighConfidence = "N";
                    }
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else if (ClsOCRValues.IsEDIDocument == "Y") // Added on 30-Dec-2021 => EDI Document
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, clsF27.clsModule.clsPickupValuesFRP001ConsigneeAddress1));
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesConsigneeAddressLine1 = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 ConsigneeCityOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesConsigneeCity = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objConsigneeCityOCR = new clsCnCWord();
                    objConsigneeCityOCR = (clsCnCWord)dataRowsShipper[0]["cncCity"];
                    PossibleWords.Add(objConsigneeCityOCR);
                    var NineDigitcount = objConsigneeCityOCR.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                    var EightDigitcount = objConsigneeCityOCR.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
                    var ZeroDigitcount = objConsigneeCityOCR.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
                    if (objConsigneeCityOCR.Confidence >= 85 && (objConsigneeCityOCR.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
                    {
                        ClsOCRValues.ConsigneeCityOCRWithHighConfidence = "Y";
                    }
                    else
                    {
                        ClsOCRValues.ConsigneeCityOCRWithHighConfidence = "N";
                    }
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else if (ClsOCRValues.IsEDIDocument == "Y") // Added on 30-Dec-2021 => EDI Document
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, clsF27.clsModule.clsPickupValuesFRP001ConsigneeCity));
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesConsigneeCity = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 ConsigneeStateOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesConsigneeState = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objConsigneeStateOCR = new clsCnCWord();
                    objConsigneeStateOCR = (clsCnCWord)dataRowsShipper[0]["cncState"];
                    PossibleWords.Add(objConsigneeStateOCR);

                    var NineDigitcount = objConsigneeStateOCR.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                    var EightDigitcount = objConsigneeStateOCR.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
                    var ZeroDigitcount = objConsigneeStateOCR.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
                    if (objConsigneeStateOCR.Confidence >= 85 && (objConsigneeStateOCR.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
                    {
                        ClsOCRValues.ConsigneeStateOCRWithHighConfidence = "Y";
                    }
                    else
                    {
                        ClsOCRValues.ConsigneeStateOCRWithHighConfidence = "N";
                    }
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else if (ClsOCRValues.IsEDIDocument == "Y") // Added on 30-Dec-2021 => EDI Document
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, clsF27.clsModule.clsPickupValuesFRP001ConsigneeState));
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesConsigneeState = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 ConsigneeZipOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesConsigneeZip = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objConsigneeZipOCR = new clsCnCWord();
                    objConsigneeZipOCR = (clsCnCWord)dataRowsShipper[0]["cncZip"];
                    PossibleWords.Add(objConsigneeZipOCR);
                    var NineDigitcount = objConsigneeZipOCR.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                    var EightDigitcount = objConsigneeZipOCR.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
                    var ZeroDigitcount = objConsigneeZipOCR.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
                    if (objConsigneeZipOCR.Confidence >= 85 && (objConsigneeZipOCR.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
                    {
                        ClsOCRValues.ConsigneeZipOCRWithHighConfidence = "Y";
                    }
                    else
                    {
                        ClsOCRValues.ConsigneeZipOCRWithHighConfidence = "N";
                    }
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else if (ClsOCRValues.IsEDIDocument == "Y") // Added on 30-Dec-2021 => EDI Document
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, clsF27.clsModule.clsPickupValuesFRP001ConsigneeZip));
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesConsigneeZip = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 ConsingeeTelNoOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objConsingeeTelNoOCR = new clsCnCWord();
                    objConsingeeTelNoOCR = (clsCnCWord)dataRowsShipper[0]["cncTelPhone"];
                    if (objConsingeeTelNoOCR.strWord != null)
                    {
                        PossibleWords.Add(objConsingeeTelNoOCR);
                        // check character confidence and avg confidence
                        var NineDigitcount = objConsingeeTelNoOCR.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                        var EightDigitcount = objConsingeeTelNoOCR.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
                        var ZeroDigitcount = objConsingeeTelNoOCR.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
                        if (objConsingeeTelNoOCR.Confidence >= 85 && (objConsingeeTelNoOCR.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
                        {
                            ClsOCRValues.ConsigneeTelephoneOCRWithHighConfidence = "Y";
                        }
                        else
                        {
                            ClsOCRValues.ConsigneeTelephoneOCRWithHighConfidence = "N";
                        }
                    }
                    else
                    {
                        PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                    }
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }

            ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        // Bill To 3rd party
        private RetStructF3 BillToNameOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesBillToName = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objBillToNameOCR = new clsCnCWord();
                    //objAdd1 = (clsCnCWord)Module1.TableRecordssFinal.Rows[0]["cncName"];
                    objBillToNameOCR = (clsCnCWord)dataRowsShipper[0]["cncName"];
                    PossibleWords.Add(objBillToNameOCR);
                    var NineDigitcount = objBillToNameOCR.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                    var EightDigitcount = objBillToNameOCR.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
                    var ZeroDigitcount = objBillToNameOCR.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
                    if (objBillToNameOCR.Confidence >= 85 && (objBillToNameOCR.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
                    {
                        ClsOCRValues.BillToNameOCRWithHighConfidence = "Y";
                    }
                    else
                    {
                        ClsOCRValues.BillToNameOCRWithHighConfidence = "N";
                    }
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesBillToName = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 BillToAddressLineOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesBillToAddressLine1 = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objBillToAddressLineOCR = new clsCnCWord();
                    objBillToAddressLineOCR = (clsCnCWord)dataRowsShipper[0]["cncAddress"];
                    PossibleWords.Add(objBillToAddressLineOCR);
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesBillToAddressLine1 = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 BillToCityOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesBillToCity = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objBillToCityOCR = new clsCnCWord();
                    objBillToCityOCR = (clsCnCWord)dataRowsShipper[0]["cncCity"];
                    PossibleWords.Add(objBillToCityOCR);
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesBillToCity = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 BillToStateOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesBillToState = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objBillToStateOCR = new clsCnCWord();
                    objBillToStateOCR = (clsCnCWord)dataRowsShipper[0]["cncState"];
                    PossibleWords.Add(objBillToStateOCR);
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesBillToState = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 BillToZipOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesBillToZip = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objBillToZipOCR = new clsCnCWord();
                    objBillToZipOCR = (clsCnCWord)dataRowsShipper[0]["cncZip"];
                    PossibleWords.Add(objBillToZipOCR);
                }
                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }
            ClsOCRValues.ClsOCRValuesBillToZip = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }
        private RetStructF3 BillToTelNoOCR(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesBillToTelePhoneNo = string.Empty;
            if (Module1.TableRecordssFinal != null && Module1.TableRecordssFinal.Rows.Count > 0)
            {
                // Check that shipper address is present in the Final table
                DataRow[] dataRowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
                if (dataRowsShipper != null && dataRowsShipper.Length > 0)
                {
                    clsCnCWord objBillToTelNoOCR = new clsCnCWord();
                    objBillToTelNoOCR = (clsCnCWord)dataRowsShipper[0]["cncTelPhone"];
                    if (objBillToTelNoOCR.strWord != null)
                    {
                        PossibleWords.Add(objBillToTelNoOCR);
                        var NineDigitcount = objBillToTelNoOCR.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                        var EightDigitcount = objBillToTelNoOCR.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
                        var ZeroDigitcount = objBillToTelNoOCR.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
                        if (objBillToTelNoOCR.Confidence >= 85 && (objBillToTelNoOCR.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
                        {
                            ClsOCRValues.BillToTelephoneOCRWithHighConfidence = "Y";
                        }
                        else
                        {
                            ClsOCRValues.BillToTelephoneOCRWithHighConfidence = "N";
                        }
                    }
                    else
                    {
                        PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                    }
                }

                else
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                }
            }
            else
            {
                PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            }

            ClsOCRValues.ClsOCRValuesBillToTelePhoneNo = PossibleWords[0].strWord;
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            return Module1.objRetStructF3;
        }


        #region Terms - Prepaid, Collect and 3rd Party
        private RetStructF3 F3_AutoLocate_Terms(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {
            //RetStructF3 objRetStructF3 = new RetStructF3();
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            try
            {
                Module1.objRetStructF3 = Terms(ObjMetaData, intCurrPageNumber, strArg);
            }
            catch (Exception ex)
            {
                //MessageBox.Show("F3: F3_AutoLocate_Terms " + ex.Message);
            }
            return Module1.objRetStructF3;

        }
        private RetStructF3 Terms(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {
            //RetStructF3 objRetStructF3 = new RetStructF3();
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            //List<int> lstLineNo = new List<int>();
            //clsCnCLine currentline = new clsCnCLine();
            //clsCnCWord objWord = new clsCnCWord();
            clsCNCSBR ObjSbr = new clsCNCSBR();
            //DataTable dt = new DataTable();
            //dt = GetDBDataForTerms(PRONo);
            //if (dt.Rows.Count > 0)
            //{
            //foreach (DataRow row2 in dt.Rows)
            //{
            string ShipperName = string.Empty;
            string ConsigneeName = string.Empty;
            string BillToName = string.Empty;
            string ShipperAddress = string.Empty;
            string ConsigneeAddress = string.Empty;
            string BillToAddress = string.Empty;
            // Temp code commented on 2-Dec-2020 for testing -- Start
            ShipperName = clsF27.clsModule.clsFinalValuesShipperName.Trim();
            ConsigneeName = clsF27.clsModule.clsFinalValuesConsigneeName.Trim();
            BillToName = clsF27.clsModule.clsFinalValuesBillToName.Trim();
            ShipperAddress = clsF27.clsModule.clsFinalValuesShipperAddress1.Trim() + clsF27.clsModule.clsFinalValuesShipperCity.Trim() + clsF27.clsModule.clsFinalValuesShipperState.Trim() + clsF27.clsModule.clsFinalValuesShipperZip.Trim();
            ConsigneeAddress = clsF27.clsModule.clsFinalValuesConsigneeAddress1.Trim() + clsF27.clsModule.clsFinalValuesConsigneeCity.Trim() + clsF27.clsModule.clsFinalValuesConsigneeState.Trim() + clsF27.clsModule.clsFinalValuesConsigneeZip.Trim();
            BillToAddress = clsF27.clsModule.clsFinalValuesBillToAddress1.Trim() + clsF27.clsModule.clsFinalValuesBillToCity.Trim() + clsF27.clsModule.clsFinalValuesBillToState.Trim() + clsF27.clsModule.clsFinalValuesBillToZip.Trim();
            // Temp code commented on 2-Dec-2020 for testing -- end

            // STEP-1(A) :- Bill-To Name found => then execute the following steps
            if (!string.IsNullOrEmpty(BillToName))
            {
                if (!string.IsNullOrEmpty(BillToName) && !string.IsNullOrEmpty(ShipperName) && !string.IsNullOrEmpty(ConsigneeName))
                {
                    int PerBillToShipperName = 0;
                    int PerBillToConsigneeName = 0;
                    if (BillToName != "*****") // check BilltoName should not be 5 star
                    {
                        PerBillToShipperName = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(BillToName), Module1.func_RemoveSpecialCharacter(ShipperName)).PercentageMatch;
                        PerBillToConsigneeName = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(BillToName), Module1.func_RemoveSpecialCharacter(ConsigneeName)).PercentageMatch;
                    }
                    // STEP-2 :- If (ShipperName && ConsineeName && BillToName) = all 3 different then TERMS => PREPAID
                    if (PerBillToShipperName < 80 && PerBillToConsigneeName < 80)
                    {
                        //PossibleWords.Add(GetDefaultWordTerms(intCurrPageNumber, "PPD"));
                        PossibleWords.Add(GetDefaultWordTermsRed(intCurrPageNumber, "PPD")); // Added on 7-July-2021 make it RED to avoid false positive case
                        Module1.objRetStructF3.Words = PossibleWords;
                        ConfLevel.Add(90);
                        Module1.objRetStructF3.ManualConfirmation = "N";
                        Module1.objRetStructF3.Status = "S";
                        Module1.objRetStructF3.Flag = "Y";
                        Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                        Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
                        //MessageBox.Show("Terms STEP-2 :- " + BillToName + ShipperName  + ConsigneeName);

                    }

                    // STEP-3 :- If(ShipperName && ConsineeName && BillTo Name) = all 3 same then execute the following steps using ShipperAddress,ConsineeAddress,BillToAddress
                    else if (PerBillToShipperName > 80 && PerBillToConsigneeName > 80)
                    {
                        int PerBillToShipperAddress = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(BillToAddress), Module1.func_RemoveSpecialCharacter(ShipperAddress)).PercentageMatch;
                        int PerBillToConsigneeAddress = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(BillToAddress), Module1.func_RemoveSpecialCharacter(ConsigneeAddress)).PercentageMatch;
                        // STEP-3_1 :- If (ShipperAddress && ConsineeAddress && BillTo Address) = all 3 Address same then TERMS => Rules there in RED colour
                        if (PerBillToShipperAddress > 80 && PerBillToConsigneeAddress > 80)
                        {
                            // TO DO Not yet confirm
                        }
                        // STEP-3_2 :- If(BillTo Address == ShipperAddress) are same then TERMS => PREPAID
                        else if (PerBillToShipperAddress > 80)
                        {
                            //PossibleWords.Add(GetDefaultWordTerms(intCurrPageNumber, "PREPAID"));
                            PossibleWords.Add(GetDefaultWordTerms(intCurrPageNumber, "PPD"));
                            Module1.objRetStructF3.Words = PossibleWords;
                            ConfLevel.Add(90);
                            Module1.objRetStructF3.ManualConfirmation = "N";
                            Module1.objRetStructF3.Status = "S";
                            Module1.objRetStructF3.Flag = "Y";
                            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
                            //MessageBox.Show("Terms STEP-3_2 " + BillToAddress + ShipperAddress);
                        }
                        // Step-3_3 :- If(BillToAddress == ConsigneeAddress) are same then TERMS => COLLECT
                        else if (PerBillToConsigneeAddress > 80)
                        {
                            //PossibleWords.Add(GetDefaultWordTerms(intCurrPageNumber, "COLLECT"));
                            PossibleWords.Add(GetDefaultWordTerms(intCurrPageNumber, "COL"));
                            Module1.objRetStructF3.Words = PossibleWords;
                            ConfLevel.Add(90);
                            Module1.objRetStructF3.ManualConfirmation = "N";
                            Module1.objRetStructF3.Status = "S";
                            Module1.objRetStructF3.Flag = "Y";
                            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
                            //MessageBox.Show("Terms STEP-3_3 " + BillToAddress + ConsigneeAddress);
                        }
                        // Step-3_4 :- If(BillToAddress && ShipperAddress && ConsigneeAddress) all 3 address different then F3: SearchKeyword : Method1 and Method2
                        else if (PerBillToShipperAddress < 80 && PerBillToConsigneeAddress < 80)
                        {
                            //Module1.objRetStructF3 = Terms_SearchKeyword(ObjMetaData, intCurrPageNumber, PRONo);
                            // TO DO Rules there in RED colour
                        }

                    }
                    // STEP-4 :- If (BillTo Name == ShipperName) are same then TERMS => PREPAID
                    else if (PerBillToShipperName > 80 && PerBillToShipperName > PerBillToConsigneeName)
                    {
                        //PossibleWords.Add(GetDefaultWordTerms(intCurrPageNumber, "PREPAID"));
                        PossibleWords.Add(GetDefaultWordTerms(intCurrPageNumber, "PPD"));
                        Module1.objRetStructF3.Words = PossibleWords;
                        ConfLevel.Add(90);
                        Module1.objRetStructF3.ManualConfirmation = "N";
                        Module1.objRetStructF3.Status = "S";
                        Module1.objRetStructF3.Flag = "Y";
                        Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                        Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
                        //MessageBox.Show("Terms STEP- 4 " + BillToName + ShipperName);
                    }
                    //  STEP-5 :- If (BillTo Name == ConsineeName) are same then TERMS => COLLECT
                    else if (PerBillToConsigneeName > 80 && PerBillToConsigneeName > PerBillToShipperName)
                    {
                        //PossibleWords.Add(GetDefaultWordTerms(intCurrPageNumber, "COLLECT"));
                        PossibleWords.Add(GetDefaultWordTerms(intCurrPageNumber, "COL"));
                        Module1.objRetStructF3.Words = PossibleWords;
                        ConfLevel.Add(90);
                        Module1.objRetStructF3.ManualConfirmation = "N";
                        Module1.objRetStructF3.Status = "S";
                        Module1.objRetStructF3.Flag = "Y";
                        Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                        Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
                        //MessageBox.Show("Terms STEP- 5 " + BillToName + ConsigneeName);

                    }
                }
            }
            // STEP-1(B) :- Bill-To Name NOT found => then execute the following steps => STEP-1(B)_1 :- check where all there words - Prepaid, collect, 3rd party are present on the same line then check it for the symbol "X" and STEP-1(B)_2 :-  check the left keyword and top keyword 
            // F3: SearchKeyword : Method1 and Method2
            // Temp below code commented
            //else
            //{
            //    Module1.objRetStructF3 = Terms_SearchKeyword(ObjMetaData, intCurrPageNumber, PRONo);
            //}

            // Step-6 :- If nothing found then serach using keyword
            if (Module1.objRetStructF3.Status == "F")
            {
                Module1.objRetStructF3 = Terms_SearchKeyword(ObjMetaData, intCurrPageNumber, strArg);
            }

            // If we dont find the terms then by default drop PPD with red color
            if (Module1.objRetStructF3.Status == "F")
            {
                PossibleWords.Add(GetDefaultWordTermsRed(intCurrPageNumber, "PPD"));
                Module1.objRetStructF3.Words = PossibleWords;
                ConfLevel.Add(90);
                Module1.objRetStructF3.ManualConfirmation = "N";
                Module1.objRetStructF3.Status = "S";
                Module1.objRetStructF3.Flag = "Y";
                Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            }

            //}
            //}
            return Module1.objRetStructF3;
        }
        // added on 23-March-2021 -- start
        private RetStructF3 Terms_OBL(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            ClsOCRValues.ClsOCRValuesTerms = string.Empty;
            try
            {
                Module1.objRetStructF3 = Terms_SearchKeyword(ObjMetaData, intCurrPageNumber, strArg);
                if (Module1.objRetStructF3.Status == "F")
                {
                    PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
                    Module1.objRetStructF3.Words = PossibleWords;
                    ConfLevel.Add(90);
                    Module1.objRetStructF3.ManualConfirmation = "N";
                    Module1.objRetStructF3.Status = "S";
                    Module1.objRetStructF3.Flag = "Y";
                    Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                    Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;

                }
                ClsOCRValues.ClsOCRValuesTerms = Module1.objRetStructF3.Words[0].strWord; // added on 24-March-2021
            }
            catch (Exception ex)
            {
                //MessageBox.Show("F3: F3_AutoLocate_Terms " + ex.Message);
            }
            return Module1.objRetStructF3;
        }
        // added on 23-March-2021 -- end
        #endregion


        private RetStructF3 AllAddresses(clsCncMetaData ObjMetaData, int intCurrPageNumber, string PRONo)
        {
            //RetStructF3 objRetStructF3 = new RetStructF3();
            Module1.TableRecordss = new DataTable();
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            ClsCNC objCNC = new ClsCNC();
            clsCNCSBR ObjSbr = new clsCNCSBR();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            List<int> lstLineNo = new List<int>();
            clsCnCLine currentline = new clsCnCLine();
            //RetStructIQSBR050 objRet50 = new RetStructIQSBR050();
            RetStructIQSBR026 objRet26 = new RetStructIQSBR026();
            string str_LastWordFromDB = string.Empty;
            string DBValue = string.Empty;
            // bool isFoundAddress = false;
            DataTable dtKeysConsignee = new DataTable();
            DataTable dtKeysShipper = new DataTable();
            DataTable dtUSCityStateZip = new DataTable();
            DataTable wordsDictionary = ObjMetaData.DocDictionary;
            //string[] arrStateCode = dtUSCityStateZip.AsEnumerable().Select(t => t.Field<string>("state")).ToArray<string>();
            Regex rxTel = new Regex(@"(\d{3}\s+\d{3}\s+\d{4}|\d{3}\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{1}\s*\-\s*\d{3}\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{1}\s*\.\s*\d{3}\s*\.\s*\d{3}\s*\.\s*\d{4}|\(\s*\d{3}\s*\)\s*\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\.\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\s*\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*\.\s*\d{3}\s*\.\s*\d{4}|\d{3}\s*\.\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*/\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*\-\s*\d{3}\s*/\s*\d{4}|\d{3}\s*/\s*\d{3}\s*/\s*\d{4}|^\d{10}$)"); //RajeshB Added 28DEC20 ^\d{10}$
            //Regex rxTel = new Regex(@"(\d{3}\s+\d{3}\s+\d{4}|\d{3}\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{1}\s*\-\s*\d{3}\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{1}\s*\.\s*\d{3}\s*\.\s*\d{3}\s*\.\s*\d{4}|\(\s*\d{3}\s*\)\s*\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\.\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\s*\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*\.\s*\d{3}\s*\.\s*\d{4}|\d{3}\s*\.\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*/\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*\-\s*\d{3}\s*/\s*\d{4}|\d{3}\s*/\s*\d{3}\s*/\s*\d{4})");
            //Regex rxTel_10digit = new Regex("^\\d{10}$");  //("^[0-9-]*$");
            Regex rxAddNumericWithTwoUpperCaseAplpha = new Regex("[0-9][A-Z]{2}");  //("^[0-9-]*$");
            Regex regex_email = new Regex(@"([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)");  //17Jun21
            //Regex rxStreetNumber = new Regex("^[0-9]+$");
            Regex rxStreetNumber = new Regex("^[0-9]+");
            string AbbrivationPattern_Endwith = "AVE|DR|RD|COURT|HWY|ST|PARK|BLVD|PKWY|EXPWY|LANE|BOX|BUILDING|CENTER|SUITE|BLDG|VILLA|HILL|FRWY|LANE|EXT|LOOP|PIKE|PLAZA|ROAD|AVENUE|PARKWAY|HIGHWAY|WAY";
            // string AbbrivationPattern_Endwith = "AVE|DR|RD|COURT|HWY|ST|PARK|BLVD|PKWY|EXPWY|LANE|BUILDING|CENTER|SUITE|BLDG|VILLA|HILL|FRWY|LANE|EXT|LOOP|PIKE|PLAZA|ROAD|AVENUE|PARKWAY|HIGHWAY";
            //string AbbrivationPattern = " AVE| DR| RD| COURT| HWY| ST| PARK| BLVD| PKWY| EXPWY| LANE| BOX| BUILDING| CENTER| SUITE| BLDG| VILLA| HILL| FRWY| LANE| EXT| LOOP| PIKE| PLAZA| ROAD| AVENUE| PARKWAY| HIGHWAY";
            string AbbrivationPattern = " AVE | DR | RD | COURT | HWY | ST | PARK | BLVD | PKWY | EXPWY | LANE | BOX | BUILDING | CENTER | SUITE | BLDG | VILLA | HILL | FRWY | LANE | EXT | LOOP | PIKE | PLAZA | ROAD | AVENUE | PARKWAY | HIGHWAY | WAY ";
            string[] Abbrivation = { "AVE", "DR", "RD", "COURT", "HWY", "ST", "PARK", "BLVD", "PKWY", "EXPWY", "LANE", "BOX", "BUILDING", "CENTER", "SUITE", "BLDG", "VILLA", "HILL", "FRWY", "LANE", "EXT", "LOOP", "PIKE", "PLAZA", "ROAD", "AVENUE", "PARKWAY", "HIGHWAY", "WAY" };
            string[] StopWordsAddress = { "Address", "Addr", "Dress" };
            string[] StopWordsName = { "Name", "Name:", "Name :" };
            string[] arrStateCode = {
                        "AA", "AE", "AP", "AK", "AL", "AR", "AZ", "CA", "CO", "CT", "DC", "DE", "FL", "GA", "GU", "HI", "IA", "ID", "IL", "IN",
                        "KS", "KY", "LA", "MA", "MD", "ME", "MI", "MN", "MO", "MS", "MT", "NC", "ND", "NE", "NH", "NJ", "NM", "NV", "NY", "OH",
                        "OK", "OR", "PA", "PR", "RI", "SC", "SD", "TN", "TX", "UT", "VA", "VI", "VT", "WA", "WI", "WV", "WY"};

            string[] arrStateCodeFullName = {
                "alabama","alaska","arizona","arkansas","california","colorado","connecticut","delaware","florida","georgia","hawaii","idaho",
                "illinois","Indiana","Iiowa","kansas","kentucky","louisiana","maine","maryland","massachusetts","Mmichigan","minnesota","mississippi",
                "missouri","montanaNebraska","nevada","new hampshire","new jersey","new mexico","new york","north carolina","north dakota","ohio",
                "oklahoma","oregon","pennsylvania","rhode island","south carolina","south dakota","tennessee","texas","utah","vermont","virginia",
                "washington","west virginia","wisconsin","wyoming"};

            string KeywordString = String.Join(",", Array.ConvertAll(arrStateCode, z => "'" + z.Replace("'", "''") + "'"));
            string KeywordString_FullName = String.Join(",", Array.ConvertAll(arrStateCodeFullName, z => "'" + z.Replace("'", "''") + "'"));
            // string USstate = "AA|AE|AP|AK|AL|AR|AZ|CA|CO|CT|DC|DE|FL|GA|GU|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|PR|RI|SC|SD|TN|TX|UT|VA|VI|VT|WA|WI|WV|WY";

            string zipPattern = "'aannnnn','aannnnnsnnnn'";//RajeshB Added 28DEC20

            //DataRow[] dataRows = wordsDictionary.Select("([word_Otext] IN (" + KeywordString + ") OR [word_Ntext] IN (" + KeywordString + ")) AND [Word_Case]= 'U'  AND [Page_no] = " + intCurrPageNumber + "");
            //DataRow[] dataRows = wordsDictionary.Select("([word_Otext] IN (" + KeywordString + ") OR [word_Ntext] IN (" + KeywordString + ") OR [Pattern] IN (" + zipPattern + ") ) AND ([Word_Case] = 'U' OR [Word_Type] = 'AN')  AND [Page_no] = " + intCurrPageNumber + "");//RajeshB Added 28DEC20
            DataRow[] dataRows = wordsDictionary.Select("(([word_Otext] IN (" + KeywordString + ") OR [word_Ntext] IN (" + KeywordString + ") OR [Pattern] IN (" + zipPattern + ") ) AND ([Word_Case] = 'U' OR [Word_Type] = 'AN')  AND [Page_no] = " + intCurrPageNumber + ") OR (([word_Otext] IN (" + KeywordString_FullName + ") OR [word_Ntext] IN (" + KeywordString_FullName + ") ) AND [Page_no] = " + intCurrPageNumber + ")");
            //string pattern_StateZipCode = string.Format(@"[{0}]+[\b\s-,]*{1}\b", "AA|AE|AP|AK|AL|AR|AZ|CA|CO|CT|DC|DE|FL|GA|GU|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|PR|RI|SC|SD|TN|TX|UT|VA|VI|VT|WA|WI|WV|WY", "[[\\d]+[\\s-]*[\\d]*]?");
            string pattern_StateZipCode = string.Format(@"[{0}]+[\b\s-,.\/]*{1}\b", "AA|AE|AP|AK|AL|AR|AZ|CA|CO|CT|DC|DE|FL|GA|GU|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|PR|RI|SC|SD|TN|TX|UT|VA|VI|VT|WA|WI|WV|WY", "[[\\d]{5}[\\s-]{0,1}[\\d]{0,4}]?");


            int lineCount = ObjMetaData.Page[intCurrPageNumber].LineCount;

            ArrayList iLineFoundCount = new ArrayList();

            #region"Main Address"

            if (dataRows != null && dataRows.Length > 0)
            {
                bool isTelPhoneFound = false;
                Module1.TableRecordss.Columns.Add("Sate", typeof(string));
                Module1.TableRecordss.Columns.Add("Zip", typeof(string));
                Module1.TableRecordss.Columns.Add("City", typeof(string));
                Module1.TableRecordss.Columns.Add("AddressLine1", typeof(string));
                Module1.TableRecordss.Columns.Add("AddressLine2", typeof(string));
                Module1.TableRecordss.Columns.Add("Name", typeof(string));
                Module1.TableRecordss.Columns.Add("TelPhone", typeof(string));
                Module1.TableRecordss.Columns.Add("AddressType", typeof(string));
                Module1.TableRecordss.Columns.Add("MergeWords", typeof(string));
                Module1.TableRecordss.Columns.Add("clcncMergeWords", typeof(clsCnCWord));

                Module1.TableRecordss.Columns.Add("cncName", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncAddress", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncCity", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncState", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncZip", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncCityStateZip", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncTelPhone", typeof(clsCnCWord));

                Module1.TableRecordss.Columns.Add("cncAddressType", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("LineNo", typeof(int));

                Module1.TableRecordssFinal = Module1.TableRecordss.Clone();


                foreach (DataRow drRows in dataRows)
                {
                    List<clsCnCWord> obj_clsCncWord_Temp = new List<clsCnCWord>();
                    List<clsCnCWord> obj_clsCncWord_TempAddress1 = new List<clsCnCWord>();
                    List<clsCnCWord> obj_clsCncWord_TempName = new List<clsCnCWord>();

                    List<clsCnCWord> obj_clsCncWord_TempAddress1_Spec = new List<clsCnCWord>();


                    string ZipValue = string.Empty;
                    string StateValue = string.Empty;
                    string CityValue = string.Empty;
                    string PhoneNoValue = string.Empty;


                    clsCnCWord TempName = new clsCnCWord();

                    clsCnCWord TempZipValue = new clsCnCWord();
                    clsCnCWord TempStateValue = new clsCnCWord();
                    clsCnCWord TempCityValue = new clsCnCWord();
                    clsCnCWord TempAddress1 = new clsCnCWord();

                    clsCnCWord TempPhoneNo = new clsCnCWord();
                    clsCnCWord TempAddressType = new clsCnCWord();

                    isTelPhoneFound = false;

                    StringBuilder StrBuilderaddressLine1 = new StringBuilder();
                    StringBuilder StrBuilderName = new StringBuilder();
                    int PageNo = Convert.ToInt32(drRows["Page_no"]);
                    int LineNo = Convert.ToInt32(drRows["Line_no"]);
                    int WordNo = Convert.ToInt32(drRows["Word_no"]);
                    DataRow row = Module1.TableRecordss.NewRow();

                    if (LineNo > 0 && WordNo > 0)
                    {
                        clsCnCWord CurrentWord = new clsCnCWord();
                        clsCnCWord PreviousWord = new clsCnCWord();
                        clsCnCWord NextWord = new clsCnCWord();
                        clsCnCWord NextPreviousWord = new clsCnCWord();
                        CurrentWord = ObjMetaData.Page[PageNo].Line[LineNo].Word[WordNo];

                        int MostLeftValue = 0;





                        #region "City State ZipCode"

                        if (CurrentWord.strWord.Contains(':') || CurrentWord.strWord.Contains('#'))
                        {
                            //MessageBox.Show("Wrong State: " + CurrentWord.strWord);
                            continue;
                        }


                        if (CurrentWord.strWord.Trim().Length >= 7 && Regex.IsMatch(CurrentWord.strWord, "[0-9]+", RegexOptions.IgnoreCase) == true)
                        {
                            // STEP 1 :- Get the state code using state code array

                            clsCnCWord _TempStateValue = ObjMetaData.Page[PageNo].Line[LineNo].Word[WordNo];
                            clsCnCWord _TempZipValue = ObjMetaData.Page[PageNo].Line[LineNo].Word[WordNo];

                            StateValue = _TempStateValue.strWord;
                            ZipValue = _TempZipValue.strWord;


                            if (StateValue != null && StateValue.Length >= 2)
                            {
                                StateValue = StateValue.Trim().Substring(0, 2);
                            }
                            if (ZipValue != null && ZipValue.Length >= 2)
                            {
                                ZipValue = ZipValue.Trim().Substring(2, ZipValue.Trim().Length - 2);
                            }

                            TempStateValue = GetDefaultWord_StateZipCode(intCurrPageNumber, _TempStateValue, StateValue);

                            TempZipValue = GetDefaultWord_StateZipCode(intCurrPageNumber, _TempZipValue, ZipValue);

                        }
                        else
                        {
                            if (ObjMetaData.Page[PageNo].Line[LineNo].WordCount > WordNo)
                            {
                                string temp_Zipcode_Merge2words = string.Empty;
                                List<clsCnCWord> obj_clsCncWord_Zipcode_Merge2words = new List<clsCnCWord>();//new

                                NextWord = ObjMetaData.Page[PageNo].Line[LineNo].Word[WordNo + 1];

                                //***********New**************
                                int iword = 1;
                                if (Regex.IsMatch(NextWord.strWord, "zip|us", RegexOptions.IgnoreCase))
                                {
                                    if (WordNo + iword < ObjMetaData.Page[PageNo].Line[LineNo].WordCount)
                                    {
                                        iword += 1;
                                        NextWord = ObjMetaData.Page[PageNo].Line[LineNo].Word[WordNo + iword];
                                    }
                                }
                                //**************************************


                                //**************Added on 05March2021****

                                if (NextWord.strWord.Contains(':') || NextWord.strWord.Contains('#'))
                                {
                                    continue;
                                }

                                //*******************************************


                                temp_Zipcode_Merge2words = NextWord.strWord;//new
                                obj_clsCncWord_Zipcode_Merge2words.Add(NextWord);//new
                                if (Regex.Replace(NextWord.strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() == "")
                                {
                                    if (WordNo + iword < ObjMetaData.Page[PageNo].Line[LineNo].WordCount)
                                    // if (WordNo + 1 < ObjMetaData.Page[PageNo].Line[LineNo].WordCount)
                                    {
                                        iword += 1;
                                        NextWord = ObjMetaData.Page[PageNo].Line[LineNo].Word[WordNo + iword];
                                        //NextWord = ObjMetaData.Page[PageNo].Line[LineNo].Word[WordNo + 2];
                                        temp_Zipcode_Merge2words = temp_Zipcode_Merge2words + NextWord.strWord;//new
                                        obj_clsCncWord_Zipcode_Merge2words.Add(NextWord);//new
                                    }
                                }

                                // STEP 2 :- Get the Zip from right side of state code
                                Regex rxZip = new Regex("^\\d{5,9}$");  //("^[0-9-]*$");
                                if (NextWord.strWord != null)
                                {
                                    //ContentsOfField = Regex.Replace(ContentsOfField, "^[^a-zA-Z0-9%]+", "");
                                    // ContentsOfField = Regex.Replace(ContentsOfField, "[^a-zA-Z0-9)]+$", "");

                                    if (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(NextWord.strWord).Trim().ToString().Replace(",", ""), rxZip.ToString()) == true)
                                    {
                                        ZipValue = NextWord.strWord.Trim();
                                        //row["Zip"] = ZipValue;
                                        //row["Sate"] = StateValue;
                                        TempZipValue = NextWord;
                                    }
                                }

                                //**********Added new*********
                                if (ZipValue == "" && obj_clsCncWord_Zipcode_Merge2words.Count == 2)
                                {
                                    if (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(temp_Zipcode_Merge2words).Trim().ToString().Replace(",", ""), rxZip.ToString()) == true)
                                    {
                                        ZipValue = temp_Zipcode_Merge2words.Trim();
                                        TempZipValue = objCNC.MergeWords(obj_clsCncWord_Zipcode_Merge2words.ToArray());
                                    }
                                }

                                obj_clsCncWord_Zipcode_Merge2words = null;//new


                                // STEP 1 :- Get the state code using state code array
                                StateValue = ObjMetaData.Page[PageNo].Line[LineNo].Word[WordNo].strWord.Trim();
                                StateValue = Regex.Replace(StateValue, "[^a-z]", "", RegexOptions.IgnoreCase).Trim();
                                //row["Sate"] = StateValue;
                                TempStateValue = ObjMetaData.Page[PageNo].Line[LineNo].Word[WordNo];

                            }
                        }

                        if (ZipValue == "" || StateValue == "" || Module1.dtUSCityStateZip.Rows.Count <= 0 || WordNo <= 1)
                        {
                            continue;
                        }

                        //**********Addred on 05March21*********

                        ZipValue = Regex.Replace(ZipValue, "^[^a-zA-Z0-9]+", "");
                        ZipValue = Regex.Replace(ZipValue, "[^a-zA-Z0-9)]+$", "");

                        if (Regex.Replace(ZipValue, "[a-z0-9-]", "", RegexOptions.IgnoreCase).Trim() != "")
                        {
                            continue;
                        }

                        //**********************************

                        string _ZipValue = Regex.Replace(ZipValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim();
                        if (_ZipValue.Trim().Length >= 5)
                        {
                            _ZipValue = _ZipValue.Trim().Substring(0, 5);
                        }
                        else
                        {
                            continue;
                        }

                        DataRow[] foundCityRows = Module1.dtUSCityStateZip.Select(" zipcode =" + _ZipValue + " AND state ='" + StateValue + "'");
                        string foundCityValue = "";
                        if (foundCityRows != null && foundCityRows.Length > 0)
                        {
                            foreach (DataRow drCity in foundCityRows)
                            {
                                foundCityValue = Convert.ToString(drCity["city"]);
                                foundCityValue = foundCityValue.Trim();
                                break;
                            }
                        }

                        bool bFoundCityDB = false;
                        int _cityWordNo = 0;
                        if (WordNo > 1)
                        {
                            string ocrCityValue = string.Empty;
                            List<clsCnCWord> obj_clsCncWord_tempCity = new List<clsCnCWord>();

                            if (foundCityValue.Trim() != "")
                            {
                                for (int iwordLoop = WordNo - 1; 0 < iwordLoop; iwordLoop--)
                                {
                                    if (foundCityValue.Split(' ').Length == obj_clsCncWord_tempCity.Count)
                                    {
                                        //***********Addred rest of words AddressLine**********
                                        if (iwordLoop > 0)
                                        {
                                            if (ObjMetaData.Page[PageNo].Line[LineNo].Word[1].strWord.Trim().ToUpper() == "PO")
                                            {
                                                for (int iwordLoop_new = iwordLoop; 0 < iwordLoop_new; iwordLoop_new--)
                                                {
                                                    obj_clsCncWord_TempAddress1_Spec.Add(ObjMetaData.Page[PageNo].Line[LineNo].Word[iwordLoop_new]);
                                                }

                                                if (obj_clsCncWord_TempAddress1_Spec.Count <= 2 || obj_clsCncWord_TempAddress1_Spec.Count > 4)
                                                {
                                                    obj_clsCncWord_TempAddress1_Spec.Clear();
                                                }
                                            }
                                        }
                                        //*******************************************************

                                        break;
                                    }
                                    clsCnCWord clsocrCityValue = ObjMetaData.Page[PageNo].Line[LineNo].Word[iwordLoop];
                                    string tempclsocrCityValue = clsocrCityValue.strWord;
                                    tempclsocrCityValue = Regex.Replace(tempclsocrCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim();

                                    if (tempclsocrCityValue != "")
                                    {
                                        //****if any lable like state came between State and City ignore that word 
                                        if (Regex.IsMatch(tempclsocrCityValue, "state", RegexOptions.IgnoreCase) == true)
                                        {
                                            continue;
                                        }
                                        //*******************************
                                        if (tempclsocrCityValue.Trim().ToLower() == "st")
                                        {
                                            clsocrCityValue.strWord = "saint";
                                        }
                                        obj_clsCncWord_tempCity.Add(clsocrCityValue);
                                        _cityWordNo = iwordLoop;
                                    }
                                }
                            }
                            if (obj_clsCncWord_tempCity.Count > 0)
                            {
                                obj_clsCncWord_tempCity.Reverse();
                                clsCnCWord objmergeCityword = objCNC.MergeWords(obj_clsCncWord_tempCity.ToArray());


                                string tempCityvalue = Regex.Replace(objmergeCityword.strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim();
                                // if (Regex.IsMatch(Regex.Replace(foundCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim(), tempCityvalue, RegexOptions.IgnoreCase) == true)
                                if (Regex.IsMatch(tempCityvalue, Regex.Replace(foundCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim(), RegexOptions.IgnoreCase) == true)
                                {
                                    objmergeCityword.strWord = foundCityValue;
                                    CityValue = objmergeCityword.strWord;
                                    TempCityValue = objmergeCityword;
                                }
                                else
                                {
                                    clsCNCSBR objmatch = new clsCNCSBR();
                                    if (objmatch.IQSBR050(Regex.Replace(foundCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim(), tempCityvalue).PercentageMatch > 70)//new
                                    {
                                        objmergeCityword.strWord = foundCityValue;
                                        CityValue = objmergeCityword.strWord;
                                        TempCityValue = objmergeCityword;
                                    }
                                    else
                                    {
                                        if (obj_clsCncWord_tempCity.Count > 1)
                                        {
                                            obj_clsCncWord_tempCity.RemoveAt(0);
                                            objmergeCityword = objCNC.MergeWords(obj_clsCncWord_tempCity.ToArray());
                                            tempCityvalue = Regex.Replace(objmergeCityword.strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim();
                                            if (objmatch.IQSBR050(Regex.Replace(foundCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim(), tempCityvalue).PercentageMatch > 70)//new
                                            {
                                                objmergeCityword.strWord = foundCityValue;
                                                CityValue = objmergeCityword.strWord;
                                                TempCityValue = objmergeCityword;
                                            }

                                        }

                                    }
                                }
                            }


                            if (CityValue.Trim() != "")
                            {
                                bFoundCityDB = true;
                            }

                            //************* CITY not found using STATE and ZIP then Check if CITY exist on DB**************

                            if (CityValue.Trim() == "")
                            {
                                string citystopwords = "zip|city/|state";
                                //string citystopwords = "zip|city|state|//|:";
                                obj_clsCncWord_tempCity.Clear();
                                int icitywordCount = 0;
                                bool bFoundcitystopwords = false;

                                int distancefromZip_State = TempZipValue.Left - TempStateValue.Right;

                                for (int iwordLoop = WordNo - 1; 0 < iwordLoop; iwordLoop--)
                                {
                                    clsCnCWord clsocrCityValue = ObjMetaData.Page[PageNo].Line[LineNo].Word[iwordLoop];
                                    string tempclsocrCityValue = clsocrCityValue.strWord;

                                    if (Regex.IsMatch(tempclsocrCityValue, citystopwords, RegexOptions.IgnoreCase) == true || (obj_clsCncWord_tempCity.Count > 0 && (tempclsocrCityValue.Trim().EndsWith(",") || tempclsocrCityValue.Trim().EndsWith(":"))))
                                    {
                                        bFoundcitystopwords = true;
                                        break;
                                    }
                                    else if (icitywordCount > 3)
                                    {
                                        break;
                                    }

                                    if (ObjMetaData.Page[PageNo].Line[LineNo].Word[iwordLoop + 1].Left - clsocrCityValue.Right < (distancefromZip_State + 50))
                                    {
                                        if (Regex.Replace(tempclsocrCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() != "")
                                        {
                                            //icitywordCount += 1;
                                            if (Regex.Replace(tempclsocrCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim().ToLower() == "st")
                                            {
                                                clsocrCityValue.strWord = "saint";
                                            }
                                            obj_clsCncWord_tempCity.Add(clsocrCityValue);
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }

                                    icitywordCount += 1;
                                }

                                if (obj_clsCncWord_tempCity.Count > 0)
                                //if (obj_clsCncWord_tempCity.Count > 0 && bFoundcitystopwords == true)
                                {
                                    obj_clsCncWord_tempCity.Reverse();

                                    TempCityValue = objCNC.MergeWords(obj_clsCncWord_tempCity.ToArray());
                                    CityValue = TempCityValue.strWord;

                                }
                                else
                                {
                                    continue;
                                }


                                CityValue = Regex.Replace(CityValue, citystopwords, "", RegexOptions.IgnoreCase);
                                CityValue = Regex.Replace(CityValue, "^[^a-zA-Z0-9]+", "");


                                if (bFoundcitystopwords == false && CityValue.Trim().ToLower() == "load" && StateValue.Trim() == "ID")
                                {
                                    CityValue = "";
                                }

                            }

                            //******************************************************************

                        }
                        else
                        {
                            continue;
                        }

                        if (ZipValue == "" || StateValue == "" || CityValue.Trim() == "")
                        {
                            continue;
                        }

                        row["Zip"] = ZipValue;
                        row["Sate"] = StateValue;
                        row["City"] = CityValue;


                        #endregion
                        //*************Added on 18May21**********
                        TempStateValue.strWord = Regex.Replace(TempStateValue.strWord, "^[^a-zA-Z0-9]+", "");
                        TempZipValue.strWord = Regex.Replace(TempZipValue.strWord, "^[^a-zA-Z0-9]+", "");
                        //*************************************
                        // Code Added on 4-Dec-2020 -- start
                        row["cncCity"] = TempCityValue;
                        row["cncState"] = TempStateValue;
                        row["cncZip"] = TempZipValue;

                        MostLeftValue = TempCityValue.Left;
                        // Code Added on 4-Dec-2020 -- end

                        clsCncMetaData ObjMetaData2 = iQPUBLIC.PublicComponents.oMetaData2;

                        #region "AddressSingleLine"



                        int _specialCharCount = 0;
                        bool foundNumericStartwith = false;
                        bool founAbbrivation = false;
                        bool bStopWordFound = false;
                        bool bAddressType = false;
                        bool bPO = false;
                        if (bFoundCityDB == true && _cityWordNo > 1)
                        {


                            if (Regex.IsMatch(ObjMetaData.Page[PageNo].Line[LineNo].Word[_cityWordNo - 1].strWord, "City/|State/|Zip", RegexOptions.IgnoreCase) == false)
                            {


                                // string addressStopLinePattern = ":|To|From|Ship|Sold|Shipping|shipper|Consign|Deliver|Destination|Payor|Invoice|Third|3rd|Party|Remit|Freight|Bill|origin";
                                List<clsCnCWord> obj_clsCncWord_TempAddressLinewise = new List<clsCnCWord>();
                                List<clsCnCWord> obj_clsCncWord_TempNameLinewise = new List<clsCnCWord>();
                                List<clsCnCWord> obj_clsCncWord_TempAddressLinewise_PO = new List<clsCnCWord>();

                                string _ShipFrom = "From|Location|shipper|Origin";//Shipping
                                string _ShipTo = "Consign|Deliver|Destination|Shipped";
                                string _BillTo = "Bill|Payment|Charge|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit";//new


                                int ZoneDiff = 0;
                                ZoneDiff = TempStateValue.Left - TempCityValue.Right;
                                if (ZoneDiff == 0)
                                {
                                    ZoneDiff = TempZipValue.Left - TempStateValue.Right;
                                }

                                ZoneDiff = ZoneDiff + 100;

                                for (int iwordLoop = _cityWordNo - 1; 0 < iwordLoop; iwordLoop--)
                                {

                                    clsCnCWord clsocrCityValue = ObjMetaData.Page[PageNo].Line[LineNo].Word[iwordLoop];
                                    string tempclsocrCityValue = clsocrCityValue.strWord;


                                    if (bStopWordFound == false || ObjMetaData.Page[PageNo].Line[LineNo].Word[iwordLoop + 1].Left - clsocrCityValue.Right > ZoneDiff)
                                    {
                                        break;
                                    }

                                    if (bStopWordFound == true)
                                    {
                                        if (Regex.IsMatch(clsocrCityValue.strWord, _ShipFrom, RegexOptions.IgnoreCase) == true)
                                        {
                                            TempAddressType = clsocrCityValue;
                                            TempAddressType.strWord = "SHIPPER";
                                            row["AddressType"] = "SHIPPER";
                                            row["cncAddressType"] = TempAddressType;
                                            bAddressType = true;
                                            break;
                                        }
                                        else if (Regex.IsMatch(clsocrCityValue.strWord, _ShipTo, RegexOptions.IgnoreCase) == true)
                                        {
                                            TempAddressType = clsocrCityValue;
                                            TempAddressType.strWord = "CONSIGNEE";
                                            row["AddressType"] = "CONSIGNEE";
                                            row["cncAddressType"] = TempAddressType;
                                            bAddressType = true;
                                            if ((Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.75)) > clsocrCityValue.Top)
                                            {
                                                ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T"; // Added on 6-July-2021 => Tag consignee only when if we found keyword or matched with Pick up DB
                                            }
                                            break;
                                        }
                                        else if (Regex.IsMatch(clsocrCityValue.strWord, _BillTo, RegexOptions.IgnoreCase) == true)
                                        {

                                            TempAddressType = clsocrCityValue;
                                            TempAddressType.strWord = "3RD PARTY";
                                            row["AddressType"] = "3RD PARTY";
                                            row["cncAddressType"] = TempAddressType;
                                            bAddressType = true;
                                            break;
                                        }
                                    }

                                    if (tempclsocrCityValue.Trim() == "|" || tempclsocrCityValue.Trim() == "-")
                                    {
                                        _specialCharCount += 1;
                                        continue;
                                    }
                                    else if (tempclsocrCityValue.Trim().ToLower().EndsWith(":") || tempclsocrCityValue.Trim().ToLower().EndsWith("to") || tempclsocrCityValue.Trim().ToLower().EndsWith("from") || tempclsocrCityValue.Trim().ToLower().EndsWith("bill"))
                                    {
                                        bStopWordFound = true;
                                        //iwordLoop = 0;
                                        // break;
                                        continue;
                                    }

                                    if (Regex.Replace(tempclsocrCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() == "")
                                    {
                                        continue;
                                    }

                                    if (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(tempclsocrCityValue).Trim(), rxStreetNumber.ToString()) == true && obj_clsCncWord_TempAddressLinewise.Count >= 2 && foundNumericStartwith == false)
                                    {
                                        foundNumericStartwith = true;
                                        //obj_clsCncWord_TempAddress1.Add(clsocrCityValue);
                                        obj_clsCncWord_TempAddressLinewise.Add(clsocrCityValue);
                                        continue;
                                    }

                                    if (Regex.IsMatch(tempclsocrCityValue, AbbrivationPattern, RegexOptions.IgnoreCase) == true)
                                    {
                                        founAbbrivation = true;
                                    }


                                    if (Regex.Replace(tempclsocrCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim().ToUpper() == "PO")
                                    {
                                        bPO = true;
                                    }

                                    if (foundNumericStartwith == false && _specialCharCount <= 1)
                                    {
                                        obj_clsCncWord_TempAddressLinewise.Add(clsocrCityValue);
                                    }
                                    else
                                    {
                                        obj_clsCncWord_TempNameLinewise.Add(clsocrCityValue);
                                        // obj_clsCncWord_TempName.Add(clsocrCityValue);
                                    }
                                }


                                // intPO = intPO - 1;
                                if (bPO == true && obj_clsCncWord_TempAddressLinewise.Count > 0 && obj_clsCncWord_TempNameLinewise.Count == 0)
                                {
                                    //obj_clsCncWord_TempAddressLinewise.Reverse();
                                    int itotalWordCount = obj_clsCncWord_TempAddressLinewise.Count - 1;
                                    //for (int iwordLoop = itotalWordCount; intPO <= iwordLoop; iwordLoop--)
                                    //{
                                    //    obj_clsCncWord_TempNameLinewise.Add(obj_clsCncWord_TempAddressLinewise[iwordLoop]);
                                    //}
                                    bPO = false;

                                    for (int iwordLoop = 0; iwordLoop <= itotalWordCount; iwordLoop++)
                                    {

                                        if (bPO == false)
                                        {
                                            obj_clsCncWord_TempAddressLinewise_PO.Add(obj_clsCncWord_TempAddressLinewise[iwordLoop]);
                                        }
                                        else
                                        {
                                            obj_clsCncWord_TempNameLinewise.Add(obj_clsCncWord_TempAddressLinewise[iwordLoop]);
                                        }

                                        if (Regex.Replace(obj_clsCncWord_TempAddressLinewise[iwordLoop].strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim().ToUpper() == "PO")
                                        {
                                            bPO = true;
                                        }

                                    }
                                    obj_clsCncWord_TempAddressLinewise.Clear();


                                }

                                if (obj_clsCncWord_TempAddressLinewise.Count > 0)
                                {
                                    obj_clsCncWord_TempAddressLinewise.Reverse();
                                    obj_clsCncWord_TempAddress1.Add(objCNC.MergeWords(obj_clsCncWord_TempAddressLinewise.ToArray()));
                                }
                                else if (obj_clsCncWord_TempAddressLinewise_PO.Count > 0)
                                {
                                    obj_clsCncWord_TempAddressLinewise_PO.Reverse();
                                    obj_clsCncWord_TempAddress1.Add(objCNC.MergeWords(obj_clsCncWord_TempAddressLinewise_PO.ToArray()));
                                }

                                obj_clsCncWord_TempAddressLinewise_PO.Clear();

                                if (obj_clsCncWord_TempNameLinewise.Count > 0)
                                {
                                    obj_clsCncWord_TempNameLinewise.Reverse();
                                    obj_clsCncWord_TempName.Add(objCNC.MergeWords(obj_clsCncWord_TempNameLinewise.ToArray()));
                                }



                                if (obj_clsCncWord_TempAddress1.Count > 0 && obj_clsCncWord_TempName.Count > 0)
                                {
                                    obj_clsCncWord_TempNameLinewise.Clear();
                                    if (LineNo > 1)
                                    {
                                        int Left = obj_clsCncWord_TempName[0].Left - 50;
                                        if (Left < 0)
                                        {
                                            Left = 0;
                                        }
                                        int Right = obj_clsCncWord_TempName[0].Right + 50;

                                        string strShipFrom = "From| At |shipper|Origin";//Shipping
                                        string strShipTo = "Consign|Ship To|Deliver|Destination|Shipped To";
                                        string strBillTo = "Bill|Payment To|Charges To|PPD|Bill To|SFBT|Third|3rd|Party|Remit|freight to";//new

                                        clsCnCLine objTempAddressTypeLine = ObjMetaData.Page[PageNo].Line[LineNo - 1];

                                        for (int iWordNo = 1; iWordNo <= objTempAddressTypeLine.WordCount; iWordNo++)
                                        {
                                            clsCnCWord tempAddressTypeWord = objTempAddressTypeLine.Word[iWordNo];

                                            if (tempAddressTypeWord.Right >= Left && tempAddressTypeWord.Left <= Right)
                                            {
                                                if (obj_clsCncWord_TempName[0].strWord.ToLower().Contains("c/o") || obj_clsCncWord_TempName[0].strWord.ToLower().Contains("%"))
                                                {
                                                    obj_clsCncWord_TempNameLinewise.Add(tempAddressTypeWord);
                                                }
                                                else
                                                {
                                                    if (Regex.IsMatch(tempAddressTypeWord.strWord, strBillTo, RegexOptions.IgnoreCase) == true)
                                                    {

                                                        TempAddressType = tempAddressTypeWord;
                                                        TempAddressType.strWord = "3RD PARTY";
                                                        row["AddressType"] = "3RD PARTY";
                                                        row["cncAddressType"] = TempAddressType;
                                                        bAddressType = true;
                                                        break;
                                                    }
                                                    else if (Regex.IsMatch(tempAddressTypeWord.strWord, strShipFrom, RegexOptions.IgnoreCase) == true)
                                                    {
                                                        TempAddressType = tempAddressTypeWord;
                                                        TempAddressType.strWord = "SHIPPER";
                                                        row["AddressType"] = "SHIPPER";
                                                        row["cncAddressType"] = TempAddressType;
                                                        bAddressType = true;
                                                        break;
                                                    }
                                                    else if (Regex.IsMatch(tempAddressTypeWord.strWord, strShipTo, RegexOptions.IgnoreCase) == true)
                                                    {
                                                        TempAddressType = tempAddressTypeWord;
                                                        TempAddressType.strWord = "CONSIGNEE";
                                                        row["AddressType"] = "CONSIGNEE";
                                                        row["cncAddressType"] = TempAddressType;
                                                        bAddressType = true;
                                                        if ((Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.75)) > tempAddressTypeWord.Top)
                                                        {
                                                            ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T"; // Added on 6-July-2021 => Tag consignee only when if we found keyword or matched with Pick up DB
                                                        }
                                                        break;
                                                    }

                                                }
                                            }
                                        }

                                        if (obj_clsCncWord_TempNameLinewise.Count > 0)
                                        {
                                            obj_clsCncWord_TempName.Add(objCNC.MergeWords(obj_clsCncWord_TempNameLinewise.ToArray()));
                                        }

                                        obj_clsCncWord_TempName.Reverse();
                                    }
                                }

                                if (bAddressType = true || _specialCharCount >= 2 || (foundNumericStartwith == true && founAbbrivation == true))
                                {
                                }
                                else
                                {
                                    obj_clsCncWord_TempAddress1.Clear();
                                    obj_clsCncWord_TempName.Clear();

                                }

                            }
                        }
                        #endregion


                        if (obj_clsCncWord_TempAddress1.Count == 0 || obj_clsCncWord_TempName.Count == 0)
                        {
                            obj_clsCncWord_TempAddress1.Clear();
                            obj_clsCncWord_TempName.Clear();

                            #region "TelPhone NOT IN USE"

                            //if (LineNo > 1)
                            //{

                            //    int Right = 0;
                            //    int Left = 0;
                            //    int iwordNo = 0;

                            //    if (TempZipValue != null)
                            //    {
                            //        if (TempZipValue.strWord != null)
                            //        {
                            //            Right = TempZipValue.Right + 100;
                            //            iwordNo = (int)TempZipValue.WordNumber;
                            //        }
                            //    }
                            //    if (Right == 0)
                            //    {
                            //        Right = TempStateValue.Right + 100;
                            //        iwordNo = (int)TempZipValue.WordNumber;
                            //    }

                            //    if (TempCityValue != null)
                            //    {
                            //        if (TempCityValue.strWord != null)
                            //        {
                            //            Left = TempCityValue.Left - 100;
                            //        }
                            //    }

                            //    if (Left == 0)
                            //    {
                            //        Left = TempStateValue.Left - 100;
                            //    }

                            //    if (Left < 0)
                            //    {
                            //        Left = 0;
                            //    }


                            //    if (iwordNo < ObjMetaData.Page[intCurrPageNumber].Line[LineNo].WordCount)
                            //    {
                            //        clsCnCWord tempwordTel = ObjMetaData.Page[intCurrPageNumber].Line[LineNo].Word[iwordNo + 1];
                            //        if (Regex.IsMatch(tempwordTel.strWord, rxTel.ToString()) == true)
                            //        {
                            //            if (iwordNo >= 1)
                            //            {
                            //                clsCnCWord tempwordTel_SID = ObjMetaData.Page[intCurrPageNumber].Line[LineNo].Word[iwordNo];
                            //                if (tempwordTel_SID.strWord.ToUpper().Contains("SID") == true)
                            //                {
                            //                    continue;
                            //                }
                            //            }

                            //            isTelPhoneFound = true;
                            //            row["TelPhone"] = tempwordTel.strWord;
                            //            TempPhoneNo = tempwordTel;
                            //            row["cncTelPhone"] = TempPhoneNo;

                            //        }
                            //    }

                            //    if (isTelPhoneFound == false)
                            //    {

                            //        for (int telLineNo = LineNo; telLineNo <= LineNo + 2; telLineNo++)
                            //        {
                            //            if (telLineNo > ObjMetaData.Page[intCurrPageNumber].LineCount)
                            //            {
                            //                break;
                            //            }

                            //            for (int iWordNo = 1; iWordNo <= ObjMetaData.Page[intCurrPageNumber].Line[telLineNo].WordCount; iWordNo++)
                            //            {
                            //                clsCnCWord tempAddressWord = ObjMetaData.Page[intCurrPageNumber].Line[telLineNo].Word[iWordNo];

                            //                if (tempAddressWord.Right >= Left && tempAddressWord.Left <= Right)
                            //                {

                            //                    //if (Regex.IsMatch(tempAddressWord.strWord, rxTel_10digit.ToString()) == true)//new
                            //                    if (Regex.IsMatch(tempAddressWord.strWord, rxTel.ToString()) == true)
                            //                    {

                            //                        if (iWordNo > 1)
                            //                        {
                            //                            clsCnCWord tempwordTel_SID = ObjMetaData.Page[intCurrPageNumber].Line[telLineNo].Word[iWordNo - 1];
                            //                            if (tempwordTel_SID.strWord.ToUpper().Contains("SID") == true)
                            //                            {
                            //                                telLineNo = LineNo + 3;
                            //                                continue;
                            //                            }
                            //                        }

                            //                        isTelPhoneFound = true;
                            //                        row["TelPhone"] = tempAddressWord.strWord;
                            //                        TempPhoneNo = tempAddressWord;
                            //                        row["cncTelPhone"] = TempPhoneNo;
                            //                        telLineNo = ObjMetaData.Page[intCurrPageNumber].LineCount + 1;
                            //                        break;
                            //                    }

                            //                }
                            //            }
                            //        }
                            //    }
                            //}


                            #endregion

                            // clsCncMetaData ObjMetaData2 = iQPUBLIC.PublicComponents.oMetaData2;

                            #region "TelPhone"

                            if (LineNo > 1)
                            {

                                int Right = 0;
                                int Left = 0;
                                int iwordNo = 0;

                                if (TempZipValue != null)
                                {
                                    if (TempZipValue.strWord != null)
                                    {
                                        Right = TempZipValue.Right + 100;
                                        iwordNo = (int)TempZipValue.WordNumber;
                                    }
                                }
                                if (Right == 0)
                                {
                                    Right = TempStateValue.Right + 100;
                                    iwordNo = (int)TempZipValue.WordNumber;
                                }

                                if (TempCityValue != null)
                                {
                                    if (TempCityValue.strWord != null)
                                    {
                                        Left = TempCityValue.Left - 100;
                                    }
                                }

                                if (Left == 0)
                                {
                                    Left = TempStateValue.Left - 100;
                                }

                                if (Left < 0)
                                {
                                    Left = 0;
                                }


                                if (iwordNo < ObjMetaData.Page[intCurrPageNumber].Line[LineNo].WordCount)
                                {
                                    clsCnCWord tempwordTel = ObjMetaData.Page[intCurrPageNumber].Line[LineNo].Word[iwordNo + 1];
                                    if (Regex.IsMatch(tempwordTel.strWord, rxTel.ToString()) == true && tempwordTel.strWord.Trim().StartsWith("0") == false && TempZipValue.Right + 300 > tempwordTel.Left)
                                    {
                                        if (iwordNo >= 1)
                                        {
                                            clsCnCWord tempwordTel_SID = ObjMetaData.Page[intCurrPageNumber].Line[LineNo].Word[iwordNo];
                                            if (tempwordTel_SID.strWord.ToUpper().Contains("SID") == true || tempwordTel_SID.strWord.ToUpper().Contains("CID") == true)
                                            {
                                                continue;
                                            }
                                        }

                                        isTelPhoneFound = true;
                                        row["TelPhone"] = tempwordTel.strWord;
                                        TempPhoneNo = tempwordTel;
                                        row["cncTelPhone"] = TempPhoneNo;

                                    }
                                }




                                if (isTelPhoneFound == false)
                                {

                                    for (int telLineNo = LineNo + 1; telLineNo <= LineNo + 3; telLineNo++)
                                    {
                                        if (telLineNo > ObjMetaData2.Page[intCurrPageNumber].LineCount)
                                        {
                                            break;
                                        }

                                        for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[telLineNo].WordCount; iWordNo++)
                                        {
                                            clsCnCWord tempAddressWord = ObjMetaData2.Page[intCurrPageNumber].Line[telLineNo].Word[iWordNo];

                                            if (tempAddressWord.Right >= Left && tempAddressWord.Left <= Right)
                                            {

                                                //**********if string have phone: 8172220690
                                                string strTelphone = tempAddressWord.strWord;

                                                if (Regex.IsMatch(strTelphone, "hone:|p:|ph:|Phn:", RegexOptions.IgnoreCase) == true)
                                                {
                                                    strTelphone = Regex.Replace(strTelphone, "[^0-9]", "", RegexOptions.IgnoreCase).Trim();
                                                }

                                                //********************
                                                Match oTelMatch = Regex.Match(strTelphone, rxTel.ToString());


                                                //if (Regex.IsMatch(tempAddressWord.strWord, rxTel_10digit.ToString()) == true)//new
                                                if (oTelMatch.Success)
                                                {

                                                    if (iWordNo > 1)
                                                    {
                                                        clsCnCWord tempwordTel_SID = ObjMetaData2.Page[intCurrPageNumber].Line[telLineNo].Word[iWordNo - 1];
                                                        if (tempwordTel_SID.strWord.ToUpper().Contains("SID") == true || tempwordTel_SID.strWord.ToUpper().Contains("CID") == true)
                                                        {
                                                            telLineNo = LineNo + 3;
                                                            continue;
                                                        }
                                                    }

                                                    clsCnCWord cncTelNo = GetDefaultWord_StateZipCode(intCurrPageNumber, tempAddressWord, oTelMatch.Value);
                                                    isTelPhoneFound = true;
                                                    row["TelPhone"] = cncTelNo.strWord;
                                                    TempPhoneNo = cncTelNo;
                                                    row["cncTelPhone"] = TempPhoneNo;
                                                    telLineNo = ObjMetaData2.Page[intCurrPageNumber].LineCount + 1;
                                                    break;
                                                }

                                            }
                                        }
                                    }
                                }
                            }


                            #endregion

                            #region"Address NOT IN USE"

                            //bool bAddressStopline = false;
                            //bool bAddressStartWithNumeric = false;

                            //if (LineNo > 1)
                            //{
                            //    //string nameStopLinePattern = "To :|Name|Ship|Sold|Location|Shipping|shipper|Consig|Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight";
                            //    string nameStopLinePattern = "To :| To |From|Name|Ship|Sold|Location|Shipping|shipper|Consign|Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight|Bill To|origin|c/o";


                            //    int Left = TempCityValue.Left - 100;
                            //    if (Left < 0)
                            //    {
                            //        Left = 0;
                            //    }
                            //    int Right = TempZipValue.Right + 300;
                            //    //int Right = TempZipValue.Right + 100;

                            //    bool bStopNameWord = false;
                            //    ArrayList iLoopCount = new ArrayList();
                            //    for (int addLineNo = LineNo - 1; 1 < addLineNo; addLineNo--)
                            //    {
                            //        int bAddressStartWithNumeric_LineNo = 0;//new
                            //        List<clsCnCWord> obj_clsCncWord_TempAddressLinewise = new List<clsCnCWord>();
                            //        for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].WordCount; iWordNo++)
                            //        {
                            //            clsCnCWord tempAddressWord = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo];

                            //            //*****new*********
                            //            if (Regex.Replace(tempAddressWord.strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() == "")
                            //            {
                            //                continue;
                            //            }
                            //            //*******************
                            //            //STEP1: AddressLine Within CityStateZip RIO
                            //            if (tempAddressWord.Right >= Left && tempAddressWord.Left <= Right)
                            //            {
                            //                if (isTelPhoneFound == false)
                            //                {
                            //                    //if (Regex.IsMatch(tempAddressWord.strWord, rxTel_10digit.ToString()) == true)
                            //                    if (Regex.IsMatch(tempAddressWord.strWord, rxTel.ToString()) == true)
                            //                    {
                            //                        isTelPhoneFound = true;
                            //                        row["TelPhone"] = tempAddressWord.strWord;
                            //                        TempPhoneNo = tempAddressWord;
                            //                        row["cncTelPhone"] = TempPhoneNo;
                            //                        break;
                            //                    }
                            //                    else if (Regex.IsMatch(tempAddressWord.strWord.Trim(), rxTel.ToString()) == true)
                            //                    //else if (Regex.IsMatch(tempAddressWord.strWord.Trim(), rxTel_10digit.ToString()) == true)
                            //                    {
                            //                        if (iWordNo > 1)
                            //                        {
                            //                            clsCnCWord tempAddressWord2 = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo + 1];

                            //                            if (Regex.IsMatch(tempAddressWord2.strWord, "tel|phone", RegexOptions.IgnoreCase) == true)
                            //                            {
                            //                                isTelPhoneFound = true;
                            //                                row["TelPhone"] = tempAddressWord.strWord;
                            //                                TempPhoneNo = tempAddressWord;
                            //                                row["cncTelPhone"] = TempPhoneNo;
                            //                                break;
                            //                            }
                            //                        }
                            //                    }
                            //                }

                            //                if (iLoopCount.Contains(tempAddressWord.LineNo) == false)
                            //                {
                            //                    iLoopCount.Add(tempAddressWord.LineNo);
                            //                }
                            //                //STEP5: 
                            //                //**********check if line have stop keyword like Ship To/ Sold To/ Shipper From/ Shipper To/ Consignee To/
                            //                if (Regex.IsMatch(tempAddressWord.strWord, nameStopLinePattern, RegexOptions.IgnoreCase))
                            //                {
                            //                    bStopNameWord = true;
                            //                    break;
                            //                }
                            //                else
                            //                {
                            //                    //if (iWordNo > 1)
                            //                    //{
                            //                    //    clsCnCWord tempstopNameword = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1];

                            //                    //    if (Regex.IsMatch(tempstopNameword.strWord, nameStopLinePattern, RegexOptions.IgnoreCase))
                            //                    //    {
                            //                    //        bStopNameWord = true;
                            //                    //        break;
                            //                    //    }
                            //                    //}
                            //                }
                            //                //****************************************************************************************************


                            //                if (bAddressStartWithNumeric == true && bAddressStopline == false && obj_clsCncWord_TempAddress1.Count > 0)
                            //                {
                            //                    if (tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                            //                    {
                            //                        bAddressStopline = true;
                            //                    }
                            //                    else
                            //                    {
                            //                        if (iWordNo > 1)
                            //                        {
                            //                            clsCnCWord tempAddressWord2 = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1];

                            //                            if (tempAddressWord2.strWord.ToLower().Contains("addre") || tempAddressWord2.strWord.ToLower().Contains("dress"))
                            //                            {
                            //                                bAddressStopline = true;
                            //                            }
                            //                        }
                            //                    }

                            //                    //int Linenumber = Convert.ToInt32(iLoopCount[0]);

                            //                    //if (bAddressStopline == true)
                            //                    //{
                            //                    //    obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                            //                    //}
                            //                    if (bAddressStopline == true || (bAddressStartWithNumeric == true && bAddressStartWithNumeric_LineNo == tempAddressWord.LineNo))//new
                            //                    {
                            //                        obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                            //                    }
                            //                    else
                            //                    {
                            //                        bAddressStopline = true;//Stop Address Line 
                            //                        break;
                            //                    }
                            //                }

                            //                else
                            //                {
                            //                    //STEP2: Word StartWith Numeric
                            //                    //if (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(tempAddressWord.strWord.Trim().Split(' ')[0]).Trim(), rxStreetNumber.ToString()) == true || Regex.IsMatch(Module1.func_RemoveSpecialCharacter(tempAddressWord.strWord.Split(' ')[tempAddressWord.strWord.Split(' ').Length - 1]).Trim(), AbbrivationPattern_Endwith) == true)
                            //                    string abc = Regex.Replace(tempAddressWord.strWord, "[^a-z0-9\\s]", "", RegexOptions.IgnoreCase).Trim();
                            //                    string tempname = " " + abc.Split(' ')[abc.Split(' ').Length - 1] + " ";
                            //                    if (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(tempAddressWord.strWord.Trim().Split(' ')[0]).Trim(), rxStreetNumber.ToString()) == true || Regex.IsMatch(tempname, AbbrivationPattern) == true)
                            //                    {
                            //                        bAddressStartWithNumeric = true;
                            //                        bAddressStartWithNumeric_LineNo = tempAddressWord.LineNo;//new
                            //                    }

                            //                    if (tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                            //                    {
                            //                        bAddressStopline = true;
                            //                    }
                            //                    else
                            //                    {
                            //                        if (iWordNo > 1)
                            //                        {
                            //                            //clsCnCWord tempAddressWord2 = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1];

                            //                            if (ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("addre") || ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("dress"))
                            //                            {
                            //                                bAddressStopline = true;
                            //                            }
                            //                        }
                            //                    }

                            //                    obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                            //                }


                            //            }

                            //        }



                            //        if (obj_clsCncWord_TempAddressLinewise.Count > 0)
                            //        {
                            //            obj_clsCncWord_TempAddress1.Add(objCNC.MergeWords(obj_clsCncWord_TempAddressLinewise.ToArray()));

                            //            //
                            //            //STEP3: If Address Keyword found on line Stop AddressLine 
                            //            if (bAddressStopline == true || obj_clsCncWord_TempAddress1.Count > 2)
                            //            {
                            //                //****************Check Before Closing Line have Address Keyword
                            //                if (addLineNo > 1 && bStopNameWord == false)
                            //                {
                            //                    addLineNo = addLineNo - 1;
                            //                    obj_clsCncWord_TempAddressLinewise = new List<clsCnCWord>();
                            //                    bAddressStopline = false;
                            //                    for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].WordCount; iWordNo++)
                            //                    {
                            //                        clsCnCWord tempAddressWord = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo];

                            //                        if (Regex.Replace(tempAddressWord.strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() == "")
                            //                        {
                            //                            continue;
                            //                        }

                            //                        if (tempAddressWord.Right >= Left && tempAddressWord.Left <= Right)
                            //                        {

                            //                            if (tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                            //                            {
                            //                                bAddressStopline = true;
                            //                            }
                            //                            else
                            //                            {
                            //                                if (iWordNo > 1)
                            //                                {

                            //                                    if (ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("addre") || ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("dress"))
                            //                                    {
                            //                                        bAddressStopline = true;
                            //                                    }
                            //                                }
                            //                            }

                            //                            if (bAddressStopline == true)
                            //                            {
                            //                                obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                            //                            }
                            //                            else
                            //                            {
                            //                                break;
                            //                            }
                            //                        }

                            //                    }

                            //                    if (obj_clsCncWord_TempAddressLinewise.Count > 0)
                            //                    {
                            //                        obj_clsCncWord_TempAddress1.Add(objCNC.MergeWords(obj_clsCncWord_TempAddressLinewise.ToArray()));
                            //                    }
                            //                }
                            //                //************************************************************


                            //                break;


                            //            }
                            //        }

                            //        if (iLoopCount.Count >= 3 || bAddressStopline == true || bStopNameWord == true)
                            //        {
                            //            break;
                            //        }
                            //    }

                            //    //***********if Address have morthen 1 line and if StopNameKeyword Found and AddressStartWithNumeric =false then remove top line from Address 
                            //    if (obj_clsCncWord_TempAddress1.Count > 1 && bAddressStartWithNumeric == false && bStopNameWord == true)
                            //    {
                            //        obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                            //    }

                            //    //**********************************************************************
                            //}



                            #endregion

                            #region "Address"

                            if (LineNo > 1)
                            {
                                //string addressStopLinePattern = "To :| To |From|Ship|Sold|Location|Shipping|shipper|Consign|Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight|Bill To|origin|Trailer #";
                                string addressStopLinePattern = "Bill To|Location|Shipping|shipper|Consign|Deliver|Destination|Trailer #|Pick up|Receiver|B/L to|To :| To |To:|From|Ship|Sold|PPD|SFBT|Payor|Invoice|Third|Party|Remit|Freight|origin";

                                bool bStartWithNumericOREndWithAbbrivationORFoundAddressLabel = false;
                                bool bStopLabelWordFound = false;
                                bool bStopNameWord = false;
                                bool bAddressWord = false;
                                int Left = TempCityValue.Left - 100;
                                if (Left < 0)
                                {
                                    Left = 0;
                                }
                                int Right = TempZipValue.Right + 100;
                                ArrayList iLoopCount = new ArrayList();
                                // bool bcheckNextLine = false;
                                int bAddressStartWithNumeric_LineNo = 0;//new
                                for (int addLineNo = LineNo - 1; 1 < addLineNo; addLineNo--)
                                {
                                    bAddressWord = false;
                                    List<clsCnCWord> obj_clsCncWord_TempAddressLinewise = new List<clsCnCWord>();
                                    for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].WordCount; iWordNo++)
                                    {
                                        clsCnCWord tempAddressWord = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo];
                                        if (Regex.Replace(tempAddressWord.strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() == "")
                                        {
                                            continue;
                                        }

                                        if (tempAddressWord.Right >= Left && tempAddressWord.Left <= Right)
                                        {
                                            //check Telphone number
                                            //if (isTelPhoneFound == false)
                                            //{
                                            if (Regex.IsMatch(tempAddressWord.strWord, rxTel.ToString()) == true && tempAddressWord.strWord.Trim().StartsWith("0") == false)
                                            {
                                                if (isTelPhoneFound == false)
                                                {
                                                    isTelPhoneFound = true;
                                                    row["TelPhone"] = tempAddressWord.strWord;
                                                    TempPhoneNo = tempAddressWord;
                                                    row["cncTelPhone"] = TempPhoneNo;
                                                }
                                                break;
                                            }
                                            //}
                                            //************************

                                            if (iLoopCount.Contains(tempAddressWord.LineNo) == false)
                                            {
                                                iLoopCount.Add(tempAddressWord.LineNo);

                                            }

                                            //*******Added on 05March2021*********
                                            if (tempAddressWord.strWord.Trim().Split(' ').Length > 10)
                                            {
                                                addLineNo = 0;
                                                break;
                                            }
                                            //*****************************************

                                            //STEP1: check StopLabelKeyWord Found like Ship To /consignee To / Bill To/   
                                            if (Regex.IsMatch(tempAddressWord.strWord, addressStopLinePattern, RegexOptions.IgnoreCase))
                                            {
                                                bStopLabelWordFound = true;
                                                if (obj_clsCncWord_TempAddress1.Count > 0)//if stop keyword found but address list have 0 count
                                                {
                                                    addLineNo = 0;
                                                    break;
                                                }
                                            }
                                            else
                                            {
                                                if (iWordNo > 1)
                                                {
                                                    clsCnCWord tempstopNameword = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1];

                                                    if (Regex.IsMatch(tempstopNameword.strWord, addressStopLinePattern, RegexOptions.IgnoreCase))
                                                    {
                                                        bStopLabelWordFound = true;
                                                        if (obj_clsCncWord_TempAddress1.Count > 0) //if stop keyword found but address list have 0 count
                                                        {
                                                            addLineNo = 0;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }

                                            if (Regex.IsMatch(tempAddressWord.strWord, "name|Narne", RegexOptions.IgnoreCase))
                                            {
                                                bStopNameWord = true;
                                                addLineNo = 0;
                                                break;
                                            }
                                            else
                                            {
                                                if (iWordNo > 1)
                                                {
                                                    clsCnCWord tempstopNameword = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1];

                                                    if (Regex.IsMatch(tempstopNameword.strWord, "name|Narne", RegexOptions.IgnoreCase))
                                                    {
                                                        bStopNameWord = true;
                                                        addLineNo = 0;
                                                        break;
                                                    }
                                                }
                                            }
                                            //bStopNameWord
                                            //*******************

                                            // STEP3:*****if bStartWithNumericOREndWithAbbrivationORFoundAddressLabel found still check on next line have address line

                                            if (bStartWithNumericOREndWithAbbrivationORFoundAddressLabel == true && bAddressStartWithNumeric_LineNo != tempAddressWord.LineNo)
                                            {
                                                if (tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                                                {
                                                    bAddressWord = true;
                                                    string straddressLine = tempAddressWord.strWord;
                                                    straddressLine = Regex.Replace(straddressLine, "address|addre|dress", "", RegexOptions.IgnoreCase);
                                                    straddressLine = Regex.Replace(straddressLine, "[^a-z0-9]", "", RegexOptions.IgnoreCase);
                                                    if (straddressLine.Trim().Length >= 2)
                                                    {
                                                        //if (Regex.Replace(tempAddressWord.strWord, "address|addre|dress", "", RegexOptions.IgnoreCase).Trim().Length >= 4)
                                                        //{
                                                        obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                                    }
                                                    continue;
                                                }
                                                else if (iWordNo > 1)
                                                {
                                                    if (ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("addre") || ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("dress"))
                                                    {
                                                        bAddressWord = true;
                                                        string straddressLine = tempAddressWord.strWord;
                                                        straddressLine = Regex.Replace(straddressLine, "address|addre|dress", "", RegexOptions.IgnoreCase);
                                                        straddressLine = Regex.Replace(straddressLine, "[^a-z0-9]", "", RegexOptions.IgnoreCase);
                                                        if (straddressLine.Trim().Length >= 2)
                                                        {
                                                            //if (Regex.Replace(tempAddressWord.strWord, "address|addre|dress", "", RegexOptions.IgnoreCase).Trim().Length >= 4)
                                                            //{
                                                            obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                                        }
                                                        continue;
                                                    }
                                                    else if (bAddressWord == true)
                                                    {
                                                        //***********Added on 16March21************
                                                        if (tempAddressWord.strWord.Trim().EndsWith(":") == true)
                                                        {
                                                            if (tempAddressWord.Right + 10 < MostLeftValue)
                                                            {
                                                                continue;
                                                            }
                                                        }
                                                        //**********************************************
                                                        obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                                        continue;
                                                    }
                                                    else
                                                    {
                                                        bStopLabelWordFound = true;
                                                        addLineNo = 0;
                                                        break;
                                                    }
                                                }
                                                else if (bAddressWord == true)
                                                {
                                                    //***********Added on 16March21************
                                                    if (tempAddressWord.strWord.Trim().EndsWith(":") == true)
                                                    {
                                                        if (tempAddressWord.Right + 10 < MostLeftValue)
                                                        {
                                                            continue;
                                                        }
                                                    }
                                                    //**********************************************
                                                    obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                                    continue;
                                                }
                                                else
                                                {
                                                    bStopLabelWordFound = true;
                                                    addLineNo = 0;
                                                    break;
                                                }
                                            }
                                            //************

                                            //STEP2: Word StartWith Numeric

                                            string abc = Regex.Replace(tempAddressWord.strWord, "[^a-z0-9\\s]", "", RegexOptions.IgnoreCase).Trim();
                                            string tempname = " " + abc.Split(' ')[abc.Split(' ').Length - 1] + " ";

                                            if (Regex.IsMatch(Regex.Replace(tempAddressWord.strWord.Trim().Split(' ')[0], "[^a-z0-9#]", "", RegexOptions.IgnoreCase).Trim(), rxStreetNumber.ToString()) == true || Regex.IsMatch(tempname, AbbrivationPattern, RegexOptions.IgnoreCase) == true || Regex.IsMatch(tempAddressWord.strWord, AbbrivationPattern, RegexOptions.IgnoreCase) == true || tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                                            // if (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(tempAddressWord.strWord.Trim().Split(' ')[0]).Trim(), rxStreetNumber.ToString()) == true || Regex.IsMatch(tempname, AbbrivationPattern, RegexOptions.IgnoreCase) == true || Regex.IsMatch(tempAddressWord.strWord, AbbrivationPattern, RegexOptions.IgnoreCase) == true || tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                                            {
                                                bStartWithNumericOREndWithAbbrivationORFoundAddressLabel = true;
                                                bAddressStartWithNumeric_LineNo = tempAddressWord.LineNo;
                                            }
                                            else
                                            {
                                                if (iWordNo > 1)
                                                {
                                                    if (ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("addre") || ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("dress"))
                                                    {
                                                        bStartWithNumericOREndWithAbbrivationORFoundAddressLabel = true;
                                                        bAddressStartWithNumeric_LineNo = tempAddressWord.LineNo;
                                                    }
                                                }
                                            }

                                            //***************Added Need to check those conditions*********
                                            if (tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                                            {
                                                string straddressLine = tempAddressWord.strWord;
                                                straddressLine = Regex.Replace(straddressLine, "address|addre|dress", "", RegexOptions.IgnoreCase);
                                                straddressLine = Regex.Replace(straddressLine, "[^a-z0-9]", "", RegexOptions.IgnoreCase);
                                                if (straddressLine.Trim().Length >= 2)
                                                {
                                                    obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                                }
                                            }
                                            else
                                            {
                                                //if (Regex.Replace(tempAddressWord.strWord, "address|addre|dress", "", RegexOptions.IgnoreCase).Trim().Length >= 5)
                                                //{

                                                //***********Added on 16March21************
                                                if (tempAddressWord.strWord.Trim().EndsWith(":") == true)
                                                {
                                                    if (tempAddressWord.Right + 10 < MostLeftValue)
                                                    {
                                                        continue;
                                                    }
                                                }
                                                //**********************************************

                                                obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                                //}
                                            }
                                        }
                                    }

                                    if (obj_clsCncWord_TempAddressLinewise.Count > 0)
                                    {
                                        //***********Added on 16March21************
                                        if (MostLeftValue > obj_clsCncWord_TempAddressLinewise[0].Left)
                                        {
                                            MostLeftValue = obj_clsCncWord_TempAddressLinewise[0].Left;
                                        }
                                        //*****************************************
                                        //*****************Added on 17Jun21 //Remove Labels and Extream Left words from Word 
                                        clsCnCWord objFinalWord = objCNC.MergeWords(obj_clsCncWord_TempAddressLinewise.ToArray());
                                        if (!regex_email.Match(objFinalWord.strWord).Success) //Line should not email id
                                        {
                                            objFinalWord.strWord = Module1.func_RemoveKeywordsFromAddress(objFinalWord.strWord);
                                            if (objFinalWord.strWord.Trim() != "")
                                            {
                                                //*******************************************
                                                obj_clsCncWord_TempAddress1.Add(objFinalWord);
                                            }
                                        }
                                    }

                                    if (iLoopCount.Count >= 3 || bStopLabelWordFound == true)
                                    {
                                        break;
                                    }
                                }


                                //***********if Address have morthen 1 line and if StopNameKeyword Found and AddressStartWithNumeric =false then remove top line from Address 
                                if (obj_clsCncWord_TempAddress1.Count > 1 && bStartWithNumericOREndWithAbbrivationORFoundAddressLabel == false && bStopLabelWordFound == true)
                                {
                                    obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                                }
                                else if (obj_clsCncWord_TempAddress1.Count > 1 && bStartWithNumericOREndWithAbbrivationORFoundAddressLabel == false && bStopLabelWordFound == false && bStopNameWord == false)
                                {

                                    obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                                    if (obj_clsCncWord_TempAddress1.Count > 1)
                                    {
                                        obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                                    }
                                    if (obj_clsCncWord_TempAddress1.Count > 1)
                                    {
                                        obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                                    }
                                    if (obj_clsCncWord_TempAddress1.Count > 1)
                                    {
                                        obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                                    }
                                }
                                //**********************************************************************

                            }

                            #endregion

                            #region"Name NOT IN USE"
                            //if (obj_clsCncWord_TempAddress1.Count > 0)
                            //{
                            //    bAddressStopline = false;

                            //    bool bNameStopLine = false;
                            //    bool bNameStopLine_FirstWord = false;
                            //    int AddressLineNo = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].LineNo;
                            //    if (AddressLineNo > 1)
                            //    {
                            //        int avgHeight = 0;
                            //        //int intZipCodeTop = TempCityValue.Top;
                            //        for (int intAddLoop = 0; intAddLoop <= obj_clsCncWord_TempAddress1.Count - 1; intAddLoop++)
                            //        {
                            //            if (intAddLoop == 0)
                            //            {
                            //                avgHeight = TempCityValue.Top - obj_clsCncWord_TempAddress1[intAddLoop].Bottom;
                            //            }
                            //            else if (obj_clsCncWord_TempAddress1[intAddLoop].Top - obj_clsCncWord_TempAddress1[intAddLoop - 1].Bottom > avgHeight)
                            //            {
                            //                avgHeight = obj_clsCncWord_TempAddress1[intAddLoop].Top - obj_clsCncWord_TempAddress1[intAddLoop - 1].Bottom;
                            //            }

                            //        }
                            //        avgHeight = avgHeight + 10;

                            //        // string nameStopLinePattern_HardStop = "To :|From| To|Location|Shipping|shipper|Consigne|Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight|Sold To|Ship To|Bill To|origin|P.O. Box";
                            //        // string nameStopLinePattern = "To :|Name|From| To|Location|Shipping|shipper|Consigne|Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight|Sold To|Ship To|Bill To|origin";
                            //        string nameStopLinePattern = "To :|From| To |Location|Shipping|shipper|Consign|Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight|Sold To|Ship To|Bill To|origin";


                            //        int Left = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Left - 100;
                            //        if (Left < 0)
                            //        {
                            //            Left = 0;
                            //        }
                            //        int Right = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Right + 300;
                            //        //int Right = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Right + 100;
                            //        //int iLoopCount = 0;

                            //        ArrayList iLoopCount = new ArrayList();

                            //        ObjMetaData2 = iQPUBLIC.PublicComponents.oMetaData2;
                            //        int nameLineHeight = 0;
                            //        for (int nameLineNo = AddressLineNo - 1; 0 < nameLineNo; nameLineNo--)
                            //        {
                            //            bAddressStopline = false;
                            //            bNameStopLine = false;
                            //            bNameStopLine_FirstWord = false;

                            //            List<clsCnCWord> obj_clsCncWord_TempAddressLinewise = new List<clsCnCWord>();
                            //            List<clsCnCWord> obj_clsCncWord_TempNameLinewise = new List<clsCnCWord>();
                            //            for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].WordCount; iWordNo++)
                            //            {
                            //                clsCnCWord tempAddressWord = ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo];

                            //                //**************new*****
                            //                if (Regex.Replace(tempAddressWord.strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() == "")
                            //                {
                            //                    continue;
                            //                }
                            //                //******************

                            //                //STEP1: AddressLine Within CityStateZip RIO
                            //                if (tempAddressWord.Right >= Left && tempAddressWord.Left <= Right)
                            //                {



                            //                    //STEP5: 


                            //                    if (nameLineNo + 1 == AddressLineNo)//Check First Line have an Address Label 
                            //                    {

                            //                        if (tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                            //                        {
                            //                            bAddressStopline = true;
                            //                        }
                            //                        else
                            //                        {
                            //                            if (iWordNo > 1)
                            //                            {
                            //                                clsCnCWord tempAddressWord2 = ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1];

                            //                                if (tempAddressWord2.strWord.ToLower().Contains("addre") || tempAddressWord2.strWord.ToLower().Contains("dress"))
                            //                                {
                            //                                    bAddressStopline = true;
                            //                                }
                            //                            }
                            //                        }

                            //                        if (bAddressStopline == true)
                            //                        {
                            //                            obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                            //                        }
                            //                        //else
                            //                        //{
                            //                        //    //  bAddressStopline = true;//Stop Address Line 
                            //                        //    // break;
                            //                        //}
                            //                    }

                            //                    if (bAddressStopline == false)
                            //                    {
                            //                        if (iLoopCount.Contains(tempAddressWord.LineNo) == false)
                            //                        {
                            //                            iLoopCount.Add(tempAddressWord.LineNo);
                            //                        }

                            //                        //if (obj_clsCncWord_TempName.Count == 0)
                            //                        //{
                            //                        if (Regex.IsMatch(tempAddressWord.strWord, nameStopLinePattern, RegexOptions.IgnoreCase))
                            //                        {
                            //                            bNameStopLine = true;
                            //                            //obj_clsCncWord_TempNameLinewise.Add(tempAddressWord);
                            //                            // break;
                            //                        }
                            //                        else
                            //                        {
                            //                            if (iWordNo > 1)
                            //                            {
                            //                                //clsCnCWord tempAddressWord2 = ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1];

                            //                                if (Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, nameStopLinePattern, RegexOptions.IgnoreCase))
                            //                                {
                            //                                    //   bNameStopLine = true;
                            //                                    bNameStopLine_FirstWord = true;
                            //                                }
                            //                            }
                            //                            //obj_clsCncWord_TempNameLinewise.Add(tempAddressWord);
                            //                        }


                            //                        if (obj_clsCncWord_TempName.Count == 0)
                            //                        {
                            //                            nameLineHeight = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Top - tempAddressWord.Bottom;
                            //                        }
                            //                        else
                            //                        {
                            //                            nameLineHeight = obj_clsCncWord_TempName[obj_clsCncWord_TempName.Count - 1].Top - tempAddressWord.Bottom;
                            //                        }

                            //                        if (nameLineHeight <= avgHeight)
                            //                        {
                            //                            if (bNameStopLine == true)
                            //                            {
                            //                                if (obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Left - 100 >= tempAddressWord.Left || obj_clsCncWord_TempName.Count == 0)
                            //                                {
                            //                                    obj_clsCncWord_TempNameLinewise.Add(tempAddressWord);
                            //                                }
                            //                            }
                            //                            else
                            //                            {
                            //                                obj_clsCncWord_TempNameLinewise.Add(tempAddressWord);
                            //                            }

                            //                        }
                            //                        else
                            //                        {
                            //                            nameLineNo = 0;
                            //                            break;
                            //                        }
                            //                        //}

                            //                        if (bNameStopLine_FirstWord == true)
                            //                        {
                            //                            bNameStopLine = true;
                            //                        }
                            //                        //else
                            //                        //{
                            //                        //    if (iWordNo > 1)
                            //                        //    {
                            //                        //        // clsCnCWord tempAddressWord2 = ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1];

                            //                        //        if (Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, nameStopLinePattern, RegexOptions.IgnoreCase))
                            //                        //        {
                            //                        //            bNameStopLine = true;
                            //                        //        }
                            //                        //    }
                            //                        //    if (bNameStopLine == true)
                            //                        //    {
                            //                        //        obj_clsCncWord_TempNameLinewise.Add(tempAddressWord);
                            //                        //    }
                            //                        //    else
                            //                        //    {
                            //                        //        nameLineNo = 0; //Line Loop should Close
                            //                        //        break;
                            //                        //    }
                            //                        //}

                            //                    }

                            //                }
                            //            }



                            //            if (obj_clsCncWord_TempAddressLinewise.Count > 0)
                            //            {
                            //                obj_clsCncWord_TempAddress1.Add(objCNC.MergeWords(obj_clsCncWord_TempAddressLinewise.ToArray()));

                            //                //STEP3: If Address Keyword found on line Stop AddressLine 
                            //            }

                            //            if (obj_clsCncWord_TempNameLinewise.Count > 0)
                            //            {
                            //                obj_clsCncWord_TempName.Add(objCNC.MergeWords(obj_clsCncWord_TempNameLinewise.ToArray()));

                            //                if (obj_clsCncWord_TempName.Count == 1 && bNameStopLine == true)
                            //                {
                            //                    break;
                            //                }

                            //                if (obj_clsCncWord_TempName.Count >= 2 && bNameStopLine == true)
                            //                {
                            //                    nameLineNo = obj_clsCncWord_TempName[0].LineNo;

                            //                    bool bRemoveNameLine = false;
                            //                    for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].WordCount; iWordNo++)
                            //                    {
                            //                        clsCnCWord tempAddressWord = ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo];
                            //                        if (Regex.Replace(tempAddressWord.strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() == "")
                            //                        {
                            //                            continue;
                            //                        }
                            //                        if (tempAddressWord.Right >= Left && tempAddressWord.Left <= Right)
                            //                        {
                            //                            if (Regex.IsMatch(tempAddressWord.strWord, ":", RegexOptions.IgnoreCase))
                            //                            {
                            //                                bRemoveNameLine = true;
                            //                                break;
                            //                            }
                            //                            else
                            //                            {
                            //                                if (iWordNo > 1)
                            //                                {
                            //                                    //clsCnCWord tempAddressWord2 = ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1];

                            //                                    if (Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, ":", RegexOptions.IgnoreCase))
                            //                                    {
                            //                                        bRemoveNameLine = true;
                            //                                        break;
                            //                                    }
                            //                                }
                            //                            }
                            //                        }
                            //                    }

                            //                    if (bRemoveNameLine == true)
                            //                    {
                            //                        obj_clsCncWord_TempName.RemoveAt(0);
                            //                    }

                            //                    break;
                            //                }

                            //                //if (obj_clsCncWord_TempName.Count >= 2 && bNameStopLine == false)
                            //                //{
                            //                //    // obj_clsCncWord_TempName.RemoveAt(0);
                            //                //    obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);
                            //                //    break;
                            //                //}
                            //            }

                            //            //*****If NameStopKeyword Lebels Found and Name not found and Address have morethan 1 line then remove topline from Address and Add it into Name
                            //            if (bNameStopLine == true && obj_clsCncWord_TempName.Count == 0)
                            //            {
                            //                if (obj_clsCncWord_TempAddress1.Count > 1)
                            //                {
                            //                    obj_clsCncWord_TempName.Add(obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1]);
                            //                    obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                            //                }
                            //            }


                            //            if (iLoopCount.Count > 2 || bNameStopLine == true)
                            //            {
                            //                break;
                            //            }


                            //        }


                            //        if (bNameStopLine_FirstWord == false && bNameStopLine == false && obj_clsCncWord_TempName.Count > 1)
                            //        {
                            //            obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);
                            //            if (obj_clsCncWord_TempName.Count > 1)
                            //            {
                            //                obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);
                            //            }
                            //        }
                            //    }
                            //}
                            #endregion

                            #region "Check Hight from Name if LineCount>1"

                            ////if (obj_clsCncWord_TempName.Count > 1)
                            ////{
                            //int maxHight = 0;
                            //if (obj_clsCncWord_TempAddress1.Count > 0)
                            //{
                            //    //maxHight = TempCityValue.Top - obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Bottom;
                            //    maxHight = TempCityValue.Top - obj_clsCncWord_TempAddress1[0].Bottom;
                            //    if (obj_clsCncWord_TempAddress1.Count > 1)
                            //    {
                            //        for (int addLineNo = 1; addLineNo <= obj_clsCncWord_TempAddress1.Count - 1; addLineNo++)
                            //        //for (int addLineNo = obj_clsCncWord_TempAddress1.Count - 2; 0 <= addLineNo; addLineNo--)
                            //        {
                            //            if (obj_clsCncWord_TempAddress1[addLineNo - 1].Top - obj_clsCncWord_TempAddress1[addLineNo].Bottom > maxHight)
                            //            {
                            //                maxHight = obj_clsCncWord_TempAddress1[addLineNo - 1].Top - obj_clsCncWord_TempAddress1[addLineNo].Bottom;
                            //            }
                            //        }
                            //    }
                            //}

                            //    //int stopLine = -1;
                            //    //for (int addLineNo = obj_clsCncWord_TempName.Count - 2; 0 <= addLineNo; addLineNo--)
                            //    //{
                            //    //    if (obj_clsCncWord_TempName[addLineNo + 1].Top - obj_clsCncWord_TempName[addLineNo].Bottom > maxHight)
                            //    //    {
                            //    //        stopLine = addLineNo;
                            //    //        break;
                            //    //    }
                            //    //}

                            //    //if (stopLine >= 0)
                            //    //{

                            //    //}


                            ////}
                            #endregion

                            #region"Name"
                            if (obj_clsCncWord_TempAddress1.Count > 0)
                            {
                                int AddressLineNo = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].LineNo;
                                if (AddressLineNo > 1)
                                {
                                    //string nameStopLinePattern = "To :|From|To:|Location|Shipping|shipper|Consign|Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight bill|Sold To|Ship To|Bill To|Bills To|origin|Shipped To|Shipped|Trailer #|Pick up|Receiver|Pickup|Billed To|PREPAID CHARGES|Freight Charge|Freight PREPAID|B/L to";
                                    string nameStopLinePattern = "Freight bill|Sold To|Ship To|Bill To|Bills To|Shipped To|Pick up|Billed To|PREPAID CHARGES|Freight Charge|Freight PREPAID|B/L to|Trailer #|Location|Shipping|shipper|Shipped|Consign|Deliver|Destination|Payor|Invoice|Receiver|Pickup|Third|Party|Remit|origin|To :|From|To:|PPD|SFBT|3rd";
                                    //string nameStopLinePattern = "To :|From| To |Location|Shipping|shipper|Consign|Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight|Sold To|Ship To|Bill To|origin|Shipped To|Shipped|Trailer #|Pick up";
                                    bool bNameStopLine = false;
                                    bool bNameFound = false;
                                    int Left = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Left - 100;
                                    if (Left < 0)
                                    {
                                        Left = 0;
                                    }
                                    //int Right = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Right + 100;

                                    int Right = 0;

                                    if (obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Right > TempZipValue.Right)
                                    {
                                        Right = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Right + 100;
                                    }
                                    else
                                    {
                                        Right = TempZipValue.Right + 100;
                                    }

                                    ArrayList iLoopCount = new ArrayList();
                                    int bNameStartWithNameKeyword_LineNo = 0;//new
                                    for (int nameLineNo = AddressLineNo - 1; 0 < nameLineNo; nameLineNo--)
                                    {
                                        if (bNameStopLine == true)
                                        {
                                            break;
                                        }
                                        List<clsCnCWord> obj_clsCncWord_TempNameLinewise = new List<clsCnCWord>();
                                        for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].WordCount; iWordNo++)
                                        {
                                            clsCnCWord tempAddressWord = ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo];


                                            //**************new*****
                                            if (Regex.Replace(tempAddressWord.strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() == "")
                                            {
                                                continue;
                                            }
                                            //******************

                                            //STEP1: AddressLine Within CityStateZip RIO
                                            if (tempAddressWord.Right >= Left && tempAddressWord.Left <= Right)
                                            {



                                                //check Telphone number
                                                //if (isTelPhoneFound == false)
                                                //{
                                                if (Regex.IsMatch(tempAddressWord.strWord, rxTel.ToString()) == true && tempAddressWord.strWord.Trim().StartsWith("0") == false)
                                                {
                                                    if (isTelPhoneFound == false)
                                                    {
                                                        isTelPhoneFound = true;
                                                        row["TelPhone"] = tempAddressWord.strWord;
                                                        TempPhoneNo = tempAddressWord;
                                                        row["cncTelPhone"] = TempPhoneNo;
                                                    }
                                                    break;
                                                }
                                                //}
                                                //************************

                                                if (iLoopCount.Contains(tempAddressWord.LineNo) == false)
                                                {
                                                    iLoopCount.Add(tempAddressWord.LineNo);
                                                }


                                                //*******Added on 05March2021*********
                                                if (tempAddressWord.strWord.Trim().Split(' ').Length > 10)
                                                {
                                                    nameLineNo = 0;
                                                    break;
                                                }
                                                //*****************************************

                                                //*STEP5:  If SOLD Keyword found Ignore that address

                                                if (Regex.IsMatch(tempAddressWord.strWord, "sold", RegexOptions.IgnoreCase))
                                                {
                                                    obj_clsCncWord_TempNameLinewise.Clear();
                                                    obj_clsCncWord_TempName.Clear();
                                                    obj_clsCncWord_TempAddress1.Clear();
                                                    nameLineNo = 0;
                                                    break;
                                                }
                                                else if (iWordNo > 1)
                                                {
                                                    if (Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, "sold", RegexOptions.IgnoreCase) && tempAddressWord.Left - ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].Right < 300)
                                                    {
                                                        obj_clsCncWord_TempNameLinewise.Clear();
                                                        obj_clsCncWord_TempName.Clear();
                                                        obj_clsCncWord_TempAddress1.Clear();
                                                        nameLineNo = 0;
                                                        break;
                                                    }
                                                }
                                                //********************************


                                                //STEP2: check Stop Keyword like Ship To / Consignee To / Bill To
                                                if (Regex.IsMatch(tempAddressWord.strWord, nameStopLinePattern, RegexOptions.IgnoreCase))
                                                {
                                                    bNameStopLine = true;

                                                    //***********Added on 16March21************
                                                    if (tempAddressWord.Right + 10 < MostLeftValue)
                                                    {
                                                        continue;
                                                    }
                                                    else
                                                    {
                                                        //nameLineNo = 0;
                                                        //break;

                                                        if (obj_clsCncWord_TempName.Count > 0)
                                                        {
                                                            if (Regex.IsMatch(tempAddressWord.strWord, "ship from|ship to|shipfrom|shipto", RegexOptions.IgnoreCase))
                                                            {
                                                                if (Regex.Replace(Regex.Replace(tempAddressWord.strWord, "ship from|ship to|shipfrom|shipto", "", RegexOptions.IgnoreCase), "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim().Length <= 2)
                                                                {
                                                                    nameLineNo = 0;
                                                                    break;
                                                                }
                                                            }
                                                            else
                                                            {
                                                                nameLineNo = 0;
                                                                break;
                                                            }
                                                        }
                                                    }

                                                }
                                                else if (iWordNo > 1)
                                                {
                                                    if (Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, nameStopLinePattern, RegexOptions.IgnoreCase) && tempAddressWord.Left - ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].Right < 300)
                                                    {
                                                        bNameStopLine = true;
                                                        //if (obj_clsCncWord_TempName.Count > 0)
                                                        //{
                                                        // iWordNo = ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].WordCount + 1;
                                                        //nameLineNo = 0;
                                                        // break;
                                                        //}
                                                    }
                                                }

                                                //*********************

                                                // STEP4:*****if bNameFound found still check on next line have Name line
                                                if (bNameFound == true && bNameStartWithNameKeyword_LineNo != tempAddressWord.LineNo)
                                                {
                                                    if (tempAddressWord.strWord.ToLower().Contains("name") || tempAddressWord.strWord.ToLower().Contains("Narne"))
                                                    {
                                                        string strnameLine = tempAddressWord.strWord;
                                                        strnameLine = Regex.Replace(strnameLine, "Narne:|Narne|name:|name", "", RegexOptions.IgnoreCase);
                                                        strnameLine = Regex.Replace(strnameLine, "[^a-z0-9]", "", RegexOptions.IgnoreCase);
                                                        if (strnameLine.Trim().Length >= 2)
                                                        {
                                                            obj_clsCncWord_TempNameLinewise.Add(tempAddressWord);
                                                        }
                                                        continue;
                                                    }
                                                    else if (iWordNo > 1)
                                                    {
                                                        if (ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("name") || ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("Narne"))
                                                        {
                                                            string strnameLine = tempAddressWord.strWord;
                                                            strnameLine = Regex.Replace(strnameLine, "Narne:|Narne|name:|name", "", RegexOptions.IgnoreCase);
                                                            strnameLine = Regex.Replace(strnameLine, "[^a-z0-9]", "", RegexOptions.IgnoreCase);
                                                            if (strnameLine.Trim().Length >= 2)
                                                            {
                                                                obj_clsCncWord_TempNameLinewise.Add(tempAddressWord);
                                                            }
                                                            continue;
                                                        }
                                                        else
                                                        {
                                                            bNameStopLine = true;
                                                            nameLineNo = 0;
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        bNameStopLine = true;
                                                        nameLineNo = 0;
                                                        break;
                                                    }
                                                }

                                                //******************************************************************************************

                                                //STEP3: check Name keyword Found 
                                                if (Regex.IsMatch(tempAddressWord.strWord, "Narne|name", RegexOptions.IgnoreCase))
                                                {
                                                    bNameStartWithNameKeyword_LineNo = tempAddressWord.LineNo;
                                                    bNameFound = true;
                                                    if (Regex.Replace(tempAddressWord.strWord, "Narne:|Narne|name:|name", "", RegexOptions.IgnoreCase).Trim().Length < 3)
                                                    {
                                                        continue;
                                                    }
                                                }
                                                else if (iWordNo > 1)
                                                {

                                                    if (Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, "name", RegexOptions.IgnoreCase) || Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, "Narne", RegexOptions.IgnoreCase))
                                                    {
                                                        bNameStartWithNameKeyword_LineNo = tempAddressWord.LineNo;
                                                        bNameFound = true;
                                                    }
                                                }


                                                //***********Added on 16March21************
                                                if (tempAddressWord.strWord.Trim().EndsWith(":") == true)
                                                {
                                                    if (tempAddressWord.Right + 10 < MostLeftValue)
                                                    {
                                                        continue;
                                                    }
                                                }
                                                //**********************************************

                                                // //**************01Jun21 if word have : then ignore from start to : ****************
                                                // tempAddressWord.strWord = tempAddressWord.strWord.Substring(tempAddressWord.strWord.IndexOf(":") + 1, tempAddressWord.strWord.Length - tempAddressWord.strWord.IndexOf(":") - 1).Trim();

                                                // if (tempAddressWord.strWord.Trim() == "")
                                                // {
                                                //     continue;
                                                // }
                                                ////************************************
                                                if (obj_clsCncWord_TempNameLinewise.Count > 0)
                                                {
                                                    if (tempAddressWord.Left - obj_clsCncWord_TempNameLinewise[obj_clsCncWord_TempNameLinewise.Count - 1].Right > (obj_clsCncWord_TempNameLinewise[obj_clsCncWord_TempNameLinewise.Count - 1].AvgCharY1 * 4))// check Gap beetween two world difference
                                                    {
                                                        break;
                                                    }
                                                }


                                                //******18 Jun21 if most left world right less than city left then world length >2
                                                if (tempAddressWord.Right < TempCityValue.Left && tempAddressWord.strWord.Trim().Length <= 2)
                                                {
                                                    continue;
                                                }
                                                obj_clsCncWord_TempNameLinewise.Add(tempAddressWord);
                                            }
                                        }

                                        if (obj_clsCncWord_TempNameLinewise.Count > 0)
                                        {
                                            ////*****Remove Labels before adding to list***********
                                            //clsCnCWord clscncNameWord = objCNC.MergeWords(obj_clsCncWord_TempNameLinewise.ToArray());
                                            //if (Regex.IsMatch(clscncNameWord.strWord.Trim(), "address:|Name:|Party:|ship to:|bill to:", RegexOptions.IgnoreCase)
                                            //{
                                            //    clscncNameWord.strWord
                                            //}
                                            //else
                                            //{
                                            //    clscncNameWord.strWord = Module1.func_RemoveKeywordsforAllName(clscncNameWord.strWord);
                                            //}
                                            //*********************************

                                            //***********Added on 16March21************
                                            if (MostLeftValue > obj_clsCncWord_TempNameLinewise[0].Left)
                                            {
                                                MostLeftValue = obj_clsCncWord_TempNameLinewise[0].Left;
                                            }
                                            //*****************************************
                                            //*************10Jun21****************
                                            clsCnCWord objFinalWord = objCNC.MergeWords(obj_clsCncWord_TempNameLinewise.ToArray());

                                            //**************11Junn21 ************
                                            if (obj_clsCncWord_TempName.Count > 0)// Name 2 line should not have more than 50 gap
                                            {
                                                if (objFinalWord.Left - obj_clsCncWord_TempName[obj_clsCncWord_TempName.Count - 1].Left > 50)
                                                {
                                                    nameLineNo = 0;
                                                    break;
                                                }
                                            }


                                            //************************************
                                            if (!regex_email.Match(objFinalWord.strWord).Success) //Line should not email id
                                            {
                                                if (!Module1.func_IgnoreKeyword(objFinalWord.strWord))
                                                {
                                                    objFinalWord.strWord = Module1.func_RemoveKeywordsFromName(objFinalWord.strWord); //Remove Labels and Extream Left words from Word 

                                                    //objFinalWord.strWord = Regex.Replace(objFinalWord.strWord.Trim(), " location#| location #| loc#| loc #", "", RegexOptions.IgnoreCase);
                                                    //objFinalWord.strWord = Regex.Replace(objFinalWord.strWord.Trim(), " location#| loc#", "", RegexOptions.IgnoreCase);

                                                    if (Regex.Replace(objFinalWord.strWord, "[^a-zA-Z]", "").Trim().Length > 1) //Line Should contain atleast 2 alpha charactes 
                                                    {
                                                        obj_clsCncWord_TempName.Add(objFinalWord);
                                                    }
                                                }
                                            }
                                            //************************************
                                            //obj_clsCncWord_TempName.Add(objCNC.MergeWords(obj_clsCncWord_TempNameLinewise.ToArray()));
                                        }
                                        if (iLoopCount.Count > 3)
                                        {
                                            break;
                                        }
                                    }


                                    if (bNameFound == true)
                                    {
                                        bNameStopLine = true;
                                    }

                                    if (bNameStopLine == false && obj_clsCncWord_TempName.Count > 1)
                                    {
                                        string strFirstName = obj_clsCncWord_TempName[obj_clsCncWord_TempName.Count - 1].strWord.Trim();
                                        if (Regex.IsMatch(strFirstName, " LLC| INC| LTD|C/O", RegexOptions.IgnoreCase) == false)
                                        {
                                            obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);

                                            if (obj_clsCncWord_TempName.Count > 1)
                                            {
                                                strFirstName = obj_clsCncWord_TempName[obj_clsCncWord_TempName.Count - 1].strWord.Trim();
                                                if (Regex.IsMatch(strFirstName, " LLC| INC| LTD|C/O", RegexOptions.IgnoreCase) == false)
                                                {
                                                    obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);

                                                    if (obj_clsCncWord_TempName.Count > 1)
                                                    {
                                                        strFirstName = obj_clsCncWord_TempName[obj_clsCncWord_TempName.Count - 1].strWord.Trim();
                                                        if (Regex.IsMatch(strFirstName, " LLC| INC| LTD|C/O", RegexOptions.IgnoreCase) == false)
                                                        {
                                                            obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);

                                                            if (obj_clsCncWord_TempName.Count > 1)
                                                            {
                                                                strFirstName = obj_clsCncWord_TempName[obj_clsCncWord_TempName.Count - 1].strWord.Trim();
                                                                if (Regex.IsMatch(strFirstName, " LLC| INC| LTD|C/O", RegexOptions.IgnoreCase) == false)
                                                                {
                                                                    obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }


                                        //if (obj_clsCncWord_TempName.Count > 1)
                                        //{
                                        //    obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);
                                        //}

                                        //if (obj_clsCncWord_TempName.Count > 1)
                                        //{
                                        //    obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);
                                        //}
                                    }

                                }
                            }
                            #endregion


                            if (obj_clsCncWord_TempAddress1.Count == 0)
                            {
                                continue;
                            }
                            obj_clsCncWord_TempAddress1.Reverse();
                            obj_clsCncWord_TempName.Reverse();

                            //*********************
                            if (obj_clsCncWord_TempAddress1_Spec.Count > 0)
                            {
                                obj_clsCncWord_TempAddress1_Spec.Reverse();
                                obj_clsCncWord_TempAddress1.Add(objCNC.MergeWords(obj_clsCncWord_TempAddress1_Spec.ToArray()));
                            }
                            //*************************
                            #region"AddressType"

                            if (obj_clsCncWord_TempAddress1.Count > 0)
                            {

                                int AddressLineNo = obj_clsCncWord_TempAddress1[0].LineNo;

                                if (AddressLineNo > 1)
                                {
                                    //string nameStopLinePattern = "Name|From| To|Location|Shipping|shipper|Consigned|Cons||Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight";

                                    string strShipFrom = "From| At |shipper|Origin|Pick up|Pickup";//Shipping
                                                                                                   //string strShipTo = "Consign|Ship To|Deliver|Destination|Shipped To|Sold To";
                                    string strShipTo = "Consign|ShipTo|Ship To|Deliver|Destination|ShippedTo|Shipped To|Receiver";
                                    // string strBillTo = "Payment To|Charges To|PPD|Bill To|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight";
                                    string strBillTo = "Payment To|Charges To|PPD|BillTo|Bill To|SFBT|Third|3rd|Party|Remit|freight to|REMITTO|REMIT TO|THIRD PARTY|FREIGHT CHARGES|BILLED TO|PREPAID CHARGE|B/L to";//new

                                    int Left = obj_clsCncWord_TempAddress1[0].Left - 100;
                                    if (Left < 0)
                                    {
                                        Left = 0;
                                    }
                                    int Right = obj_clsCncWord_TempAddress1[0].Right + 300;
                                    //int iLoopCount = 0;
                                    ArrayList iLoopCount = new ArrayList();
                                    int y2Top = obj_clsCncWord_TempAddress1[0].Top;

                                    ObjMetaData2 = iQPUBLIC.PublicComponents.oMetaData2;

                                    for (int nameLineNo = AddressLineNo; 0 < nameLineNo; nameLineNo--)
                                    {

                                        if (iLoopCount.Count > 3)
                                        {
                                            break;
                                        }

                                        for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].WordCount; iWordNo++)
                                        {

                                            clsCnCWord objtempWord = (clsCnCWord)ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo].DeepClone();

                                            //STEP1: AddressLine Within CityStateZip RIO
                                            if (objtempWord.Right >= Left && objtempWord.Left <= Right && objtempWord.strWord.Split(' ').Length < 10 && y2Top - objtempWord.Bottom < 300)
                                            {

                                                //iLoopCount += 1;

                                                if (iLoopCount.Contains(objtempWord.LineNo) == false)
                                                {
                                                    iLoopCount.Add(objtempWord.LineNo);
                                                }

                                                //else
                                                //{

                                                if (Regex.IsMatch(objtempWord.strWord, strBillTo, RegexOptions.IgnoreCase) == true)
                                                {

                                                    TempAddressType = objtempWord;
                                                    TempAddressType.strWord = "3RD PARTY";
                                                    row["AddressType"] = "3RD PARTY";
                                                    row["cncAddressType"] = TempAddressType;
                                                    nameLineNo = 0;
                                                    break;
                                                }
                                                else if (Regex.IsMatch(objtempWord.strWord, strShipFrom, RegexOptions.IgnoreCase) == true)
                                                {
                                                    TempAddressType = objtempWord;
                                                    TempAddressType.strWord = "SHIPPER";
                                                    row["AddressType"] = "SHIPPER";
                                                    row["cncAddressType"] = TempAddressType;
                                                    nameLineNo = 0;
                                                    break;
                                                }
                                                else if (Regex.IsMatch(objtempWord.strWord, strShipTo, RegexOptions.IgnoreCase) == true)
                                                {
                                                    TempAddressType = objtempWord;
                                                    TempAddressType.strWord = "CONSIGNEE";
                                                    row["AddressType"] = "CONSIGNEE";
                                                    row["cncAddressType"] = TempAddressType;
                                                    nameLineNo = 0;
                                                    if ((Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.75)) > objtempWord.Top)
                                                    {
                                                        ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T"; // Added on 6-July-2021 => Tag consignee only when if we found keyword or matched with Pick up DB
                                                    }
                                                    break;
                                                }

                                                else if (iWordNo > 1)
                                                {
                                                    clsCnCWord objtempWord2 = (clsCnCWord)ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].DeepClone();

                                                    //string DataLabel = objtempWord2.strWord.Trim().Split(' ')[objtempWord2.strWord.Trim().Split(' ').Length - 1];
                                                    if (objtempWord2.strWord.Split(' ').Length < 10 && objtempWord.Left - objtempWord2.Left <= 500)
                                                    {
                                                        string DataLabel = objtempWord2.strWord;
                                                        if (Regex.IsMatch(DataLabel, strShipFrom, RegexOptions.IgnoreCase) == true)
                                                        {
                                                            TempAddressType = objtempWord2;
                                                            TempAddressType.strWord = "SHIPPER";
                                                            row["AddressType"] = "SHIPPER";
                                                            row["cncAddressType"] = TempAddressType;
                                                            nameLineNo = 0;
                                                            break;
                                                        }
                                                        else if (Regex.IsMatch(DataLabel, strShipTo, RegexOptions.IgnoreCase) == true)
                                                        {
                                                            TempAddressType = objtempWord2;
                                                            TempAddressType.strWord = "CONSIGNEE";
                                                            row["AddressType"] = "CONSIGNEE";
                                                            row["cncAddressType"] = TempAddressType;
                                                            nameLineNo = 0;
                                                            if ((Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.75)) > objtempWord2.Top)
                                                            {
                                                                ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T"; // Added on 6-July-2021 => Tag consignee only when if we found keyword or matched with Pick up DB
                                                            }
                                                            break;
                                                        }
                                                        else if (Regex.IsMatch(DataLabel, strBillTo, RegexOptions.IgnoreCase) == true)
                                                        {

                                                            TempAddressType = objtempWord2;
                                                            TempAddressType.strWord = "3RD PARTY";
                                                            row["AddressType"] = "3RD PARTY";
                                                            row["cncAddressType"] = TempAddressType;
                                                            nameLineNo = 0;
                                                            break;
                                                        }
                                                    }
                                                }

                                                // }
                                            }
                                        }
                                    }
                                    //


                                }
                            }
                            #endregion

                        }



                        #region "Check Hight from Name if LineCount>1"

                        if (obj_clsCncWord_TempName.Count > 1)
                        {
                            int maxHight = 0;

                            maxHight = TempCityValue.Top - obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Bottom;

                            if (obj_clsCncWord_TempAddress1.Count > 1)
                            {
                                //for (int addLineNo = 1; addLineNo <= obj_clsCncWord_TempAddress1.Count - 1; addLineNo++)
                                for (int addLineNo = obj_clsCncWord_TempAddress1.Count - 2; 0 <= addLineNo; addLineNo--)
                                {
                                    if (obj_clsCncWord_TempAddress1[addLineNo + 1].Top - obj_clsCncWord_TempAddress1[addLineNo].Bottom > maxHight)
                                    {
                                        maxHight = obj_clsCncWord_TempAddress1[addLineNo + 1].Top - obj_clsCncWord_TempAddress1[addLineNo].Bottom;
                                    }
                                }
                            }


                            int stopLine = -1;
                            for (int addLineNo = obj_clsCncWord_TempName.Count - 2; 0 <= addLineNo; addLineNo--)
                            {
                                if (obj_clsCncWord_TempName[addLineNo + 1].Top - obj_clsCncWord_TempName[addLineNo].Bottom > maxHight + 10)
                                {
                                    stopLine = addLineNo;
                                    break;
                                }
                            }

                            if (stopLine >= 0)
                            {
                                if (stopLine == 0)
                                {
                                    obj_clsCncWord_TempName.RemoveAt(0);
                                }
                                else if (stopLine == 1)
                                {
                                    obj_clsCncWord_TempName.RemoveAt(0);
                                    obj_clsCncWord_TempName.RemoveAt(0);

                                }
                                else if (stopLine == 2)
                                {
                                    obj_clsCncWord_TempName.RemoveAt(0);
                                    obj_clsCncWord_TempName.RemoveAt(0);
                                    obj_clsCncWord_TempName.RemoveAt(0);
                                }

                            }


                        }
                        #endregion


                        #region "remove telphone number if it above on name"
                        if (isTelPhoneFound == true && obj_clsCncWord_TempName.Count > 0)
                        {
                            if (obj_clsCncWord_TempName[0].LineNo > TempPhoneNo.LineNo)
                            {
                                isTelPhoneFound = false;
                                row["TelPhone"] = "";
                                TempPhoneNo = null;
                                row["cncTelPhone"] = TempPhoneNo;
                            }
                        }
                        #endregion 

                        //**************Added on 05March21****
                        if (objCNC.MergeWords(obj_clsCncWord_TempAddress1.ToArray()).strWord.Trim().Length <= 5)
                        {
                            continue;
                        }
                        //**************************************


                        // Temp code commented and added below on 28-Nov-2020 -- start
                        //row["AddressLine1"] = StrBuilderaddressLine1;
                        //row["Name"] = StrBuilderName;
                        if (obj_clsCncWord_TempAddress1.Count > 0)
                        {
                            row["AddressLine1"] = objCNC.MergeWords(obj_clsCncWord_TempAddress1.ToArray()).strWord;
                        }

                        //if (obj_clsCncWord_TempName.Count > 0)
                        //{
                        //    row["Name"] = objCNC.MergeWords(obj_clsCncWord_TempName.ToArray()).strWord;
                        //}
                        // Temp code commented and added below on 28-Nov-2020 -- end 

                        if (obj_clsCncWord_TempAddress1.Count > 0)
                        {
                            obj_clsCncWord_Temp.AddRange(obj_clsCncWord_TempAddress1);
                        }

                        obj_clsCncWord_Temp.Add(TempCityValue);
                        obj_clsCncWord_Temp.Add(TempStateValue);
                        //obj_clsCncWord_Temp.Add(TempZipValue);//commanted by RajeshB on 28DEC20

                        //***********RajeshB Added 28DEC20
                        //if (TempStateValue.strWord.Trim().Length <= 4)
                        //{
                        obj_clsCncWord_Temp.Add(TempZipValue);
                        //}
                        //********************************

                        row["MergeWords"] = objCNC.MergeWords(obj_clsCncWord_Temp.ToArray()).strWord;
                        row["clcncMergeWords"] = objCNC.MergeWords(obj_clsCncWord_Temp.ToArray());

                        List<clsCnCWord> objCityStateZip = new List<clsCnCWord>();
                        objCityStateZip.Add(TempCityValue);
                        objCityStateZip.Add(TempStateValue);
                        //objCityStateZip.Add(TempZipValue);//commanted by RajeshB on 28DEC20

                        //***********RajeshB Added 28DEC20
                        //if (TempStateValue.strWord.Trim().Length <= 4)
                        //{
                        objCityStateZip.Add(TempZipValue);
                        //}

                        //if (obj_clsCncWord_TempName.Count > 0)
                        //{
                        //    row["cncName"] = objCNC.MergeWords(obj_clsCncWord_TempName.ToArray());
                        //}
                        //if (obj_clsCncWord_TempName.Count > 0)
                        //{
                        //    row["Name"] = objCNC.MergeWords(obj_clsCncWord_TempName.ToArray()).strWord;
                        //}
                        if (obj_clsCncWord_TempName.Count == 0)
                        {
                            row["cncName"] = GetDefaultWord(intCurrPageNumber, "*****");
                            row["Name"] = "*****";
                        } // above if part is added on 7-jan-2021 
                        else if (obj_clsCncWord_TempName.Count > 0)
                        {
                            clsCnCWord cncName = objCNC.MergeWords(obj_clsCncWord_TempName.ToArray());
                            string strcncName = cncName.strWord;
                            if (strcncName.Split(' ').Length > 15)
                            {
                                // MessageBox.Show("NameLenght>10: " + cncName.strWord);
                                //continue;
                                row["cncName"] = GetDefaultWord(intCurrPageNumber, "*****");
                                row["Name"] = "*****";
                            }
                            else
                            {
                                //*********************If name have special char and numeric then return ****** only
                                //string RemoveWordsName = "3RD PARTY BILL|3RD PARTY BILL FREIGHT PREPAID|BILL OR REMIT|REMIT|BILL FREIGHT CHARGES|BILL THIRD PARTY INFORMATION|FREIGHT PAYMELT PLAN|FREIGHT CHARGE TERMS (FREIGHT CHARGES ARE PREPAID|FREIGHT CHARGES BILL";
                                strcncName = Module1.func_RemoveKeywordsforAllName(strcncName);
                                strcncName = Regex.Replace(strcncName.Trim(), "^[^a-zA-Z0-9]+", "").Trim(); //Remove Special Keyword from Startof string
                                //Regex rxStreetNumber = new Regex("^[0-9]+$");
                                int TotCharCount = strcncName.Replace(" ", "").Length;
                                int TotAlphaNumericCount = strcncName.Replace(" ", "").Count(char.IsLetterOrDigit);
                                int TotRemainingSpecialCharCount = TotCharCount - TotAlphaNumericCount;
                                int TotNumericCount = strcncName.Replace(" ", "").Count(char.IsDigit);
                                // Check name len should not be zero or it should be greater than 2 alphabet
                                string abc = Regex.Replace(strcncName, "[^a-z0-9\\s]", "", RegexOptions.IgnoreCase).Trim();
                                string tempname = " " + abc.Split(' ')[abc.Split(' ').Length - 1] + " ";
                                if (Regex.Replace(strcncName, "[^a-z]", "", RegexOptions.IgnoreCase).Trim().Length == 0 || Regex.Replace(strcncName, "[^a-z]", "", RegexOptions.IgnoreCase).Trim().Length <= 2)
                                {
                                    row["cncName"] = GetDefaultWord(intCurrPageNumber, "*****");
                                    row["Name"] = "*****";
                                }
                                // if 5 or more than 5 special characters are present then don't consider as valid name 
                                //else if (TotRemainingSpecialCharCount >= 5)
                                else if (TotRemainingSpecialCharCount >= 10) //change to 10 ex : H.E.B GROCERY #224 (S.A. 25) have 7 specila but it is correct name
                                {
                                    row["cncName"] = GetDefaultWord(intCurrPageNumber, "*****");
                                    row["Name"] = "*****";
                                }
                                else
                                {
                                    row["cncName"] = cncName;
                                    row["Name"] = cncName.strWord;
                                }

                            }
                        }

                        if (obj_clsCncWord_TempAddress1.Count == 0)
                        {
                            continue;
                        }
                        else if (obj_clsCncWord_TempAddress1.Count > 0)
                        {

                            clsCnCWord cncAddr = objCNC.MergeWords(obj_clsCncWord_TempAddress1.ToArray());
                            if (cncAddr.strWord.Split(' ').Length > 25 || cncAddr.strWord.Trim() == "")
                            {
                                // MessageBox.Show("AddreLenght>20: " + cncAddr.strWord);
                                continue;
                            }
                            row["cncAddress"] = cncAddr;
                            row["LineNo"] = obj_clsCncWord_TempAddress1[0].LineNo;
                        }

                        if (objCityStateZip.Count > 0)
                        {
                            row["cncCityStateZip"] = objCNC.MergeWords(objCityStateZip.ToArray());
                        }


                        if (iLineFoundCount.Contains(TempStateValue.LineNo) == false)
                        {
                            iLineFoundCount.Add(TempStateValue.LineNo);
                        }

                        Module1.TableRecordss.Rows.Add(row);
                    }
                }
                Module1.TableRecordss.DefaultView.Sort = "LineNo";
                Module1.TableRecordss = Module1.TableRecordss.DefaultView.ToTable();
            }
            #endregion


            #region  "NEW LOGIC 14JAN21- if State not Found using above logic"

            if (dataRows == null || dataRows.Length == 0)
            {

                Module1.TableRecordss.Columns.Add("Sate", typeof(string));
                Module1.TableRecordss.Columns.Add("Zip", typeof(string));
                Module1.TableRecordss.Columns.Add("City", typeof(string));
                Module1.TableRecordss.Columns.Add("AddressLine1", typeof(string));
                Module1.TableRecordss.Columns.Add("AddressLine2", typeof(string));
                Module1.TableRecordss.Columns.Add("Name", typeof(string));
                Module1.TableRecordss.Columns.Add("TelPhone", typeof(string));
                Module1.TableRecordss.Columns.Add("AddressType", typeof(string));
                Module1.TableRecordss.Columns.Add("MergeWords", typeof(string));
                Module1.TableRecordss.Columns.Add("clcncMergeWords", typeof(clsCnCWord));

                Module1.TableRecordss.Columns.Add("cncName", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncAddress", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncCity", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncState", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncZip", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncCityStateZip", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("cncTelPhone", typeof(clsCnCWord));

                Module1.TableRecordss.Columns.Add("cncAddressType", typeof(clsCnCWord));
                Module1.TableRecordss.Columns.Add("LineNo", typeof(int));

                Module1.TableRecordssFinal = Module1.TableRecordss.Clone();
            }

            //bool isTelPhoneFound = false;
            for (int iLine = 1; iLine <= lineCount; iLine++)
            {
                if (iLineFoundCount.Contains(iLine) == true)
                {
                    continue;
                }

                bool isTelPhoneFound = false;
                clsCnCLine objLine = ObjMetaData.Page[intCurrPageNumber].Line[iLine];
                string textToSearch = objLine.strLine;
                MatchCollection isMatchCollection = Regex.Matches(textToSearch, pattern_StateZipCode, RegexOptions.IgnoreCase);
                if (isMatchCollection.Count > 0)
                {
                    int ActualCityValueLength = 0;
                    for (int imatch = 0; imatch < isMatchCollection.Count; imatch++)
                    {
                        string strX2 = "";
                        ActualCityValueLength = 0;

                        strX2 = objLine.LineX2Char.Split(',')[isMatchCollection[imatch].Index + isMatchCollection[imatch].Length - 1];
                        if (strX2 == "0")
                        {
                            strX2 = objLine.LineX2Char.Split(',')[isMatchCollection[imatch].Index + isMatchCollection[imatch].Length - 2];
                        }


                        int ZipCodeRight = Convert.ToInt32(strX2);


                        List<clsCnCWord> obj_clsCncWord_Temp = new List<clsCnCWord>();
                        List<clsCnCWord> obj_clsCncWord_TempAddress1 = new List<clsCnCWord>();
                        List<clsCnCWord> obj_clsCncWord_TempName = new List<clsCnCWord>();

                        string ZipValue = string.Empty;
                        string StateValue = string.Empty;
                        string CityValue = string.Empty;
                        string PhoneNoValue = string.Empty;


                        clsCnCWord TempName = new clsCnCWord();

                        clsCnCWord TempZipValue = new clsCnCWord();
                        clsCnCWord TempStateValue = new clsCnCWord();
                        clsCnCWord TempCityValue = new clsCnCWord();
                        clsCnCWord TempAddress1 = new clsCnCWord();

                        clsCnCWord TempPhoneNo = new clsCnCWord();
                        clsCnCWord TempAddressType = new clsCnCWord();

                        isTelPhoneFound = false;

                        StringBuilder StrBuilderaddressLine1 = new StringBuilder();
                        StringBuilder StrBuilderName = new StringBuilder();
                        int PageNo = intCurrPageNumber;//Convert.ToInt32(drRows["Page_no"]);
                        int LineNo = objLine.LineNo;//Convert.ToInt32(drRows["Line_no"]);
                                                    //int WordNo = Convert.ToInt32(drRows["Word_no"]);
                        DataRow row = Module1.TableRecordss.NewRow();



                        string tempCityLine = "";
                        string strMatchValue = "";

                        strMatchValue = isMatchCollection[imatch].Value;

                        if (strMatchValue.Contains(':') || strMatchValue.Contains('#'))
                        {
                            continue;
                        }

                        if (imatch == 0)
                        {
                            tempCityLine = textToSearch.Substring(0, isMatchCollection[imatch].Index);
                        }
                        else
                        {
                            tempCityLine = textToSearch.Substring(isMatchCollection[imatch - 1].Index + isMatchCollection[imatch - 1].Length, isMatchCollection[imatch].Index - (isMatchCollection[imatch - 1].Index + isMatchCollection[imatch - 1].Length));
                        }

                        if (Regex.Replace(tempCityLine, "[^a-z]", "", RegexOptions.IgnoreCase).Trim().Length <= 2)
                        {
                            continue;
                        }



                        strMatchValue = Regex.Replace(strMatchValue, "^[^a-zA-Z0-9]+", "").Trim();
                        StateValue = strMatchValue.Substring(0, 2).Trim();
                        ZipValue = strMatchValue.Substring(2, strMatchValue.Length - 2);
                        ZipValue = Regex.Replace(ZipValue, "^[^a-zA-Z0-9]+", "").Trim();




                        if (ZipValue == "" || StateValue == "" || Module1.dtUSCityStateZip.Rows.Count <= 0)
                        {
                            continue;
                        }

                        if (ZipValue.Trim().Length >= 5)
                        {
                            ZipValue = ZipValue.Trim().Substring(0, 5);
                        }
                        else
                        {
                            continue;
                        }

                        if (Regex.Replace(strMatchValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim().Length < 7 || Regex.Replace(StateValue, "[^a-z]", "", RegexOptions.IgnoreCase).Trim() == "" || Regex.Replace(ZipValue, "[^0-9]", "", RegexOptions.IgnoreCase).Trim().Length < 5)
                        {
                            continue;
                        }


                        TempStateValue = GetDefaultWord_StateZipCode_NewLogic(intCurrPageNumber, objLine, StateValue);

                        TempZipValue = GetDefaultWord_StateZipCode_NewLogic(intCurrPageNumber, objLine, ZipValue);

                        #region "City"

                        string foundCityValue = "";
                        try
                        {
                            DataRow[] foundCityRows = Module1.dtUSCityStateZip.Select(" zipcode =" + ZipValue + " AND state ='" + StateValue + "'");

                            if (foundCityRows != null && foundCityRows.Length > 0)
                            {
                                foreach (DataRow drCity in foundCityRows)
                                {
                                    foundCityValue = Convert.ToString(drCity["city"]);
                                    foundCityValue = foundCityValue.Trim();
                                    break;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show(ex.Message);
                            foundCityValue = "";
                        }

                        bool bCityFoundDB = false;
                        int icityWordCount = 0;
                        if (tempCityLine.Trim().Length > 0)
                        {
                            string ocrCityValue = string.Empty;
                            List<clsCnCWord> obj_clsCncWord_tempCity = new List<clsCnCWord>();
                            string strCityValue = "";
                            string[] arTempCityWords = tempCityLine.Trim().Split(' ');

                            if (foundCityValue.Trim() != "")
                            {

                                icityWordCount = 0;
                                for (int iwordLoop = arTempCityWords.Length - 1; 0 <= iwordLoop; iwordLoop--)
                                {
                                    string tempclsocrCityValue = arTempCityWords[iwordLoop];

                                    if (foundCityValue.Split(' ').Length == icityWordCount)
                                    {
                                        break;
                                    }
                                    //else if (Regex.IsMatch(tempclsocrCityValue, citystopwords, RegexOptions.IgnoreCase) == true)
                                    //{
                                    //    bFoundcitystopwords = true;
                                    //    break;
                                    //}
                                    ActualCityValueLength = ActualCityValueLength + tempclsocrCityValue.Length + 1;

                                    if (Regex.Replace(tempclsocrCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() != "")
                                    {
                                        //****if any lable like state came between State and City ignore that word 
                                        if (Regex.Replace(tempclsocrCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim().ToLower() == "state")
                                        {
                                            continue;
                                        }
                                        //*******************************

                                        if (Regex.Replace(tempclsocrCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim().ToLower() == "st")
                                        {
                                            tempclsocrCityValue = "saint";
                                        }
                                        strCityValue = tempclsocrCityValue + " " + strCityValue;
                                        icityWordCount += 1;
                                    }
                                }
                            }

                            if (strCityValue.Trim().Length > 0)
                            {


                                string tempCityvalue = Regex.Replace(strCityValue.Trim(), "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim();
                                //if (Regex.IsMatch(Regex.Replace(foundCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim(),tempCityvalue, RegexOptions.IgnoreCase) == true)
                                if (Regex.IsMatch(tempCityvalue, Regex.Replace(foundCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim(), RegexOptions.IgnoreCase) == true)
                                {
                                    // objmergeCityword.strWord = foundCityValue;
                                    CityValue = foundCityValue;
                                    //TempCityValue = objmergeCityword;
                                }
                                else
                                {
                                    clsCNCSBR objmatch = new clsCNCSBR();
                                    if (objmatch.IQSBR050(Regex.Replace(foundCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim(), tempCityvalue).PercentageMatch > 70)//new
                                    {
                                        //objmergeCityword.strWord = foundCityValue;
                                        CityValue = foundCityValue;
                                        //TempCityValue = objmergeCityword;
                                    }
                                }
                            }


                            if (CityValue.Trim() != "")
                            {
                                bCityFoundDB = true;
                            }

                            //************** CITY not found using STATE and ZIP then Check if CITY exist on DB***************

                            if (CityValue.Trim() == "")
                            {
                                string citystopwords = "zip|city/|state";
                                ActualCityValueLength = 0;
                                icityWordCount = 0;

                                for (int iwordLoop = arTempCityWords.Length - 1; 0 <= iwordLoop; iwordLoop--)
                                {
                                    string tempclsocrCityValue = arTempCityWords[iwordLoop];

                                    if (Regex.IsMatch(tempclsocrCityValue, citystopwords, RegexOptions.IgnoreCase) == true || (obj_clsCncWord_tempCity.Count > 0 && (tempclsocrCityValue.Trim().EndsWith(",") || tempclsocrCityValue.Trim().EndsWith(":"))))
                                    {
                                        CityValue = strCityValue.Trim();
                                        break;
                                    }

                                    //*******Added on 05March2021
                                    if (icityWordCount > 3)
                                    {
                                        CityValue = "";
                                        break;
                                    }
                                    //**************

                                    ActualCityValueLength = ActualCityValueLength + tempclsocrCityValue.Length + 1;
                                    if (Regex.Replace(tempclsocrCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() != "")
                                    {
                                        if (Regex.Replace(tempclsocrCityValue, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim().ToLower() == "st")
                                        {
                                            tempclsocrCityValue = "saint";
                                        }
                                        strCityValue = tempclsocrCityValue + " " + strCityValue;
                                        icityWordCount += 1;
                                    }
                                }


                                if (CityValue.Trim() != "")////*******Added on 05March2021
                                {
                                    CityValue = Regex.Replace(CityValue, citystopwords, "", RegexOptions.IgnoreCase);
                                    CityValue = Regex.Replace(CityValue, "^[^a-zA-Z0-9]+", "");
                                }

                            }

                        }



                        #endregion


                        if (ZipValue.Trim() == "" || StateValue.Trim() == "" || CityValue.Trim() == "")
                        {
                            continue;
                        }

                        TempCityValue = GetDefaultWord_StateZipCode_NewLogic(intCurrPageNumber, objLine, CityValue);

                        ActualCityValueLength = isMatchCollection[imatch].Index - ActualCityValueLength;

                        if (ActualCityValueLength < 0)
                        {
                            ActualCityValueLength = 0;
                        }

                        string strX1 = objLine.LineX1Char.Split(',')[ActualCityValueLength];
                        if (strX1 == "0")
                        {
                            strX1 = objLine.LineX2Char.Split(',')[ActualCityValueLength + 1];
                        }

                        int CityLeft = Convert.ToInt32(strX1);

                        row["Zip"] = ZipValue;
                        row["Sate"] = StateValue;
                        row["City"] = CityValue;

                        row["cncCity"] = TempCityValue;
                        row["cncState"] = TempStateValue;
                        row["cncZip"] = TempZipValue;





                        clsCncMetaData ObjMetaData2 = iQPUBLIC.PublicComponents.oMetaData2;

                        #region "TelPhone"

                        if (LineNo > 1)
                        {

                            int Right = 0;
                            int Left = 0;
                            // int iwordNo = 0;

                            if (TempZipValue != null)
                            {
                                if (TempZipValue.strWord != null)
                                {
                                    Right = ZipCodeRight + 100;
                                    //iwordNo = (int)TempZipValue.WordNumber;
                                }
                            }
                            //if (Right == 0)
                            //{
                            //    Right = TempStateValue.Right + 100;
                            //   // iwordNo = (int)TempZipValue.WordNumber;
                            //}

                            if (TempCityValue != null)
                            {
                                if (TempCityValue.strWord != null)
                                {
                                    Left = CityLeft - 100;
                                }
                            }

                            //if (Left == 0)
                            //{
                            //    Left = TempStateValue.Left - 100;
                            //}

                            if (Left < 0)
                            {
                                Left = 0;
                            }


                            //if (isTelPhoneFound == false)
                            //{

                            for (int telLineNo = LineNo + 1; telLineNo <= LineNo + 3; telLineNo++)
                            {
                                if (telLineNo > ObjMetaData2.Page[intCurrPageNumber].LineCount)
                                {
                                    break;
                                }

                                for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[telLineNo].WordCount; iWordNo++)
                                {
                                    clsCnCWord tempAddressWord = ObjMetaData2.Page[intCurrPageNumber].Line[telLineNo].Word[iWordNo];

                                    if (tempAddressWord.Right >= Left && tempAddressWord.Left <= Right)
                                    {


                                        //**********if string have phone: 8172220690
                                        string strTelphone = tempAddressWord.strWord;

                                        //if (Regex.IsMatch(strTelphone, "hone:|p:", RegexOptions.IgnoreCase) == true)
                                        if (Regex.IsMatch(strTelphone, "hone:|p:|ph:|Phn:", RegexOptions.IgnoreCase) == true)
                                        {
                                            strTelphone = Regex.Replace(strTelphone, "[^0-9]", "", RegexOptions.IgnoreCase).Trim();
                                        }

                                        //********************
                                        Match oTelMatch = Regex.Match(strTelphone, rxTel.ToString());
                                        //Match oTelMatch = Regex.Match(tempAddressWord.strWord, rxTel.ToString());


                                        //if (Regex.IsMatch(tempAddressWord.strWord, rxTel_10digit.ToString()) == true)//new
                                        if (oTelMatch.Success)
                                        {

                                            if (iWordNo > 1)
                                            {
                                                clsCnCWord tempwordTel_SID = ObjMetaData2.Page[intCurrPageNumber].Line[telLineNo].Word[iWordNo - 1];
                                                if (tempwordTel_SID.strWord.ToUpper().Contains("SID") == true || tempwordTel_SID.strWord.ToUpper().Contains("CID") == true)
                                                {
                                                    telLineNo = LineNo + 3;
                                                    continue;
                                                }
                                            }

                                            clsCnCWord cncTelNo = GetDefaultWord_StateZipCode(intCurrPageNumber, tempAddressWord, oTelMatch.Value);
                                            isTelPhoneFound = true;
                                            row["TelPhone"] = cncTelNo.strWord;
                                            TempPhoneNo = cncTelNo;
                                            row["cncTelPhone"] = TempPhoneNo;
                                            telLineNo = ObjMetaData2.Page[intCurrPageNumber].LineCount + 1;
                                            break;
                                        }

                                    }
                                }
                            }
                            // }
                        }


                        #endregion

                        #region "Address"

                        if (LineNo > 1)
                        {
                            //string addressStopLinePattern = "To :| To |From|Ship|Sold|Location|Shipping|shipper|Consign|Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight|Bill To|origin|Trailer #";
                            //string addressStopLinePattern = "To :| To |To:|From|Ship|Sold|Location|Shipping|shipper|Consign|Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|Party|Remit|Freight|Bill To|origin|Trailer #|Pick up|Receiver";
                            string addressStopLinePattern = "Bill To|Location|Shipping|shipper|Consign|Deliver|Destination|Trailer #|Pick up|Receiver|B/L to|To :| To |To:|From|Ship|Sold|PPD|SFBT|Payor|Invoice|Third|Party|Remit|Freight|origin";

                            bool bStartWithNumericOREndWithAbbrivationORFoundAddressLabel = false;
                            bool bStopLabelWordFound = false;
                            bool bStopNameWord = false;
                            bool bAddressWord = false;
                            //int Left = TempCityValue.Left - 100;
                            //if (Left < 0)
                            //{
                            //    Left = 0;
                            //}
                            //int Right = TempZipValue.Right + 100;

                            int Left = CityLeft - 100;
                            if (Left < 0)
                            {
                                Left = 0;
                            }
                            int Right = ZipCodeRight + 300;

                            ArrayList iLoopCount = new ArrayList();
                            // bool bcheckNextLine = false;
                            int bAddressStartWithNumeric_LineNo = 0;//new
                            for (int addLineNo = LineNo - 1; 1 < addLineNo; addLineNo--)
                            {
                                bAddressWord = false;
                                List<clsCnCWord> obj_clsCncWord_TempAddressLinewise = new List<clsCnCWord>();
                                for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].WordCount; iWordNo++)
                                {
                                    clsCnCWord tempAddressWord = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo];
                                    if (Regex.Replace(tempAddressWord.strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() == "")
                                    {
                                        continue;
                                    }

                                    if (tempAddressWord.Right >= Left && tempAddressWord.Left <= Right)
                                    {
                                        //check Telphone number
                                        //if (isTelPhoneFound == false)
                                        //{
                                        if (Regex.IsMatch(tempAddressWord.strWord, rxTel.ToString()) == true && tempAddressWord.strWord.Trim().StartsWith("0") == false)
                                        {
                                            if (isTelPhoneFound == false)
                                            {
                                                isTelPhoneFound = true;
                                                row["TelPhone"] = tempAddressWord.strWord;
                                                TempPhoneNo = tempAddressWord;
                                                row["cncTelPhone"] = TempPhoneNo;
                                            }
                                            break;
                                        }
                                        //}
                                        //************************

                                        if (iLoopCount.Contains(tempAddressWord.LineNo) == false)
                                        {
                                            iLoopCount.Add(tempAddressWord.LineNo);
                                        }

                                        //*******Added on 05March2021*********
                                        if (tempAddressWord.strWord.Trim().Split(' ').Length > 10)
                                        {
                                            addLineNo = 0;
                                            break;
                                        }
                                        //*****************************************

                                        //STEP1: check StopLabelKeyWord Found like Ship To /consignee To / Bill To/   
                                        if (Regex.IsMatch(tempAddressWord.strWord, addressStopLinePattern, RegexOptions.IgnoreCase))
                                        {
                                            bStopLabelWordFound = true;
                                            if (obj_clsCncWord_TempAddress1.Count > 0)
                                            {
                                                addLineNo = 0;
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            if (iWordNo > 1)
                                            {
                                                clsCnCWord tempstopNameword = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1];

                                                if (Regex.IsMatch(tempstopNameword.strWord, addressStopLinePattern, RegexOptions.IgnoreCase))
                                                {
                                                    bStopLabelWordFound = true;
                                                    if (obj_clsCncWord_TempAddress1.Count > 0)
                                                    {
                                                        addLineNo = 0;
                                                        break;
                                                    }
                                                }
                                            }
                                        }

                                        if (Regex.IsMatch(tempAddressWord.strWord, "Narne|name", RegexOptions.IgnoreCase))
                                        {
                                            bStopNameWord = true;
                                            addLineNo = 0;
                                            break;
                                        }
                                        else
                                        {
                                            if (iWordNo > 1)
                                            {
                                                clsCnCWord tempstopNameword = ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1];

                                                if (Regex.IsMatch(tempstopNameword.strWord, "Narne|name", RegexOptions.IgnoreCase))
                                                {
                                                    bStopNameWord = true;
                                                    addLineNo = 0;
                                                    break;
                                                }
                                            }
                                        }
                                        //bStopNameWord
                                        //*******************

                                        // STEP3:*****if bStartWithNumericOREndWithAbbrivationORFoundAddressLabel found still check on next line have address line

                                        if (bStartWithNumericOREndWithAbbrivationORFoundAddressLabel == true && bAddressStartWithNumeric_LineNo != tempAddressWord.LineNo)
                                        {
                                            if (tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                                            {
                                                bAddressWord = true;
                                                string straddressLine = tempAddressWord.strWord;
                                                straddressLine = Regex.Replace(straddressLine, "address|addre|dress", "", RegexOptions.IgnoreCase);
                                                straddressLine = Regex.Replace(straddressLine, "[^a-z0-9]", "", RegexOptions.IgnoreCase);
                                                if (straddressLine.Trim().Length >= 2)
                                                {
                                                    //if (Regex.Replace(tempAddressWord.strWord, "address|addre|dress", "", RegexOptions.IgnoreCase).Trim().Length >= 4)
                                                    //{
                                                    obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                                }
                                                continue;
                                            }
                                            else if (iWordNo > 1)
                                            {
                                                if (ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("addre") || ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("dress"))
                                                {
                                                    bAddressWord = true;
                                                    string straddressLine = tempAddressWord.strWord;
                                                    straddressLine = Regex.Replace(straddressLine, "address|addre|dress", "", RegexOptions.IgnoreCase);
                                                    straddressLine = Regex.Replace(straddressLine, "[^a-z0-9]", "", RegexOptions.IgnoreCase);
                                                    if (straddressLine.Trim().Length >= 2)
                                                    {
                                                        //if (Regex.Replace(tempAddressWord.strWord, "address|addre|dress", "", RegexOptions.IgnoreCase).Trim().Length >= 4)
                                                        //{
                                                        obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                                    }
                                                    continue;
                                                }
                                                else if (bAddressWord == true)
                                                {
                                                    obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                                    continue;
                                                }
                                                else
                                                {
                                                    bStopLabelWordFound = true;
                                                    addLineNo = 0;
                                                    break;
                                                }
                                            }
                                            else if (bAddressWord == true)
                                            {
                                                obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                                continue;
                                            }
                                            else
                                            {
                                                bStopLabelWordFound = true;
                                                addLineNo = 0;
                                                break;
                                            }
                                        }
                                        //************

                                        //STEP2: Word StartWith Numeric

                                        string abc = Regex.Replace(tempAddressWord.strWord, "[^a-z0-9\\s]", "", RegexOptions.IgnoreCase).Trim();
                                        string tempname = " " + abc.Split(' ')[abc.Split(' ').Length - 1] + " ";
                                        //if (Regex.IsMatch(Regex.Replace(tempAddressWord.strWord.Trim().Split(' ')[0], "[^a-z0-9#]", "", RegexOptions.IgnoreCase).Trim(), rxStreetNumber.ToString()) == true || Regex.IsMatch(tempname, AbbrivationPattern, RegexOptions.IgnoreCase) == true || Regex.IsMatch(tempAddressWord.strWord, AbbrivationPattern, RegexOptions.IgnoreCase) == true || tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                                        if (Regex.IsMatch(Regex.Replace(tempAddressWord.strWord.Trim().Split(' ')[0], "[^a-z0-9#]", "", RegexOptions.IgnoreCase).Trim(), rxStreetNumber.ToString()) == true || Regex.IsMatch(tempname, AbbrivationPattern, RegexOptions.IgnoreCase) == true || Regex.IsMatch(tempAddressWord.strWord, AbbrivationPattern, RegexOptions.IgnoreCase) == true || tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                                        // if (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(tempAddressWord.strWord.Trim().Split(' ')[0]).Trim(), rxStreetNumber.ToString()) == true || Regex.IsMatch(tempname, AbbrivationPattern, RegexOptions.IgnoreCase) == true || Regex.IsMatch(tempAddressWord.strWord, AbbrivationPattern, RegexOptions.IgnoreCase) == true || tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                                        {
                                            bStartWithNumericOREndWithAbbrivationORFoundAddressLabel = true;
                                            bAddressStartWithNumeric_LineNo = tempAddressWord.LineNo;
                                        }
                                        else
                                        {
                                            if (iWordNo > 1)
                                            {
                                                if (ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("addre") || ObjMetaData2.Page[intCurrPageNumber].Line[addLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("dress"))
                                                {
                                                    bStartWithNumericOREndWithAbbrivationORFoundAddressLabel = true;
                                                    bAddressStartWithNumeric_LineNo = tempAddressWord.LineNo;
                                                }
                                            }
                                        }

                                        //***************Added Need to check those conditions*********
                                        if (tempAddressWord.strWord.ToLower().Contains("addre") || tempAddressWord.strWord.ToLower().Contains("dress"))
                                        {
                                            string straddressLine = tempAddressWord.strWord;
                                            straddressLine = Regex.Replace(straddressLine, "address|addre|dress", "", RegexOptions.IgnoreCase);
                                            straddressLine = Regex.Replace(straddressLine, "[^a-z0-9]", "", RegexOptions.IgnoreCase);
                                            if (straddressLine.Trim().Length >= 2)
                                            {
                                                obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                            }
                                        }
                                        else
                                        {
                                            //if (Regex.Replace(tempAddressWord.strWord, "address|addre|dress", "", RegexOptions.IgnoreCase).Trim().Length >= 5)
                                            //{
                                            obj_clsCncWord_TempAddressLinewise.Add(tempAddressWord);
                                            //}
                                        }
                                    }
                                }

                                if (obj_clsCncWord_TempAddressLinewise.Count > 0)
                                {
                                    //*****************Added on 17Jun21 //Remove Labels and Extream Left words from Word 
                                    clsCnCWord objFinalWord = objCNC.MergeWords(obj_clsCncWord_TempAddressLinewise.ToArray());

                                    if (!regex_email.Match(objFinalWord.strWord).Success) //Line should not email id
                                    {
                                        objFinalWord.strWord = Module1.func_RemoveKeywordsFromAddress(objFinalWord.strWord);

                                        if (objFinalWord.strWord.Trim() != "")
                                        {
                                            //*******************************************

                                            obj_clsCncWord_TempAddress1.Add(objFinalWord);
                                        }
                                    }
                                }

                                if (iLoopCount.Count >= 3 || bStopLabelWordFound == true)
                                {
                                    break;
                                }
                            }


                            //***********if Address have morthen 1 line and if StopNameKeyword Found and AddressStartWithNumeric =false then remove top line from Address 
                            if (obj_clsCncWord_TempAddress1.Count > 1 && bStartWithNumericOREndWithAbbrivationORFoundAddressLabel == false && bStopLabelWordFound == true)
                            {
                                obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                            }
                            else if (obj_clsCncWord_TempAddress1.Count > 1 && bStartWithNumericOREndWithAbbrivationORFoundAddressLabel == false && bStopLabelWordFound == false && bStopNameWord == false)
                            {

                                obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                                if (obj_clsCncWord_TempAddress1.Count > 1)
                                {
                                    obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                                }
                                if (obj_clsCncWord_TempAddress1.Count > 1)
                                {
                                    obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                                }
                                if (obj_clsCncWord_TempAddress1.Count > 1)
                                {
                                    obj_clsCncWord_TempAddress1.RemoveAt(obj_clsCncWord_TempAddress1.Count - 1);
                                }
                            }
                            //**********************************************************************

                        }

                        #endregion

                        #region"Name"
                        if (obj_clsCncWord_TempAddress1.Count > 0)
                        {
                            int AddressLineNo = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].LineNo;
                            if (AddressLineNo > 1)
                            {
                                //string nameStopLinePattern = "To :|From|To:|Location|Shipping|shipper|Consign|Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight|Sold To|Ship To|Bill To|origin|Shipped To|Shipped|Trailer #|Pickup|Pick up|Billed To|PREPAID CHARGE|Freight Charge|Freight PREPAID";
                                string nameStopLinePattern = "Freight bill|Sold To|Ship To|Bill To|Bills To|Shipped To|Pick up|Billed To|PREPAID CHARGES|Freight Charge|Freight PREPAID|B/L to|Trailer #|Location|Shipping|shipper|Shipped|Consign|Deliver|Destination|Payor|Invoice|Receiver|Pickup|Third|Party|Remit|origin|To :|From|To:|PPD|SFBT|3rd";

                                bool bNameStopLine = false;
                                bool bNameFound = false;
                                int Left = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Left - 100;
                                if (Left < 0)
                                {
                                    Left = 0;
                                }
                                int Right = obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Right + 100;
                                ArrayList iLoopCount = new ArrayList();
                                int bNameStartWithNameKeyword_LineNo = 0;//new
                                for (int nameLineNo = AddressLineNo - 1; 0 < nameLineNo; nameLineNo--)
                                {
                                    if (bNameStopLine == true)
                                    {
                                        break;
                                    }
                                    List<clsCnCWord> obj_clsCncWord_TempNameLinewise = new List<clsCnCWord>();
                                    for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].WordCount; iWordNo++)
                                    {
                                        clsCnCWord tempAddressWord = ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo];



                                        //**************new*****
                                        if (Regex.Replace(tempAddressWord.strWord, "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim() == "")
                                        {
                                            continue;
                                        }
                                        //******************

                                        //STEP1: AddressLine Within CityStateZip RIO
                                        if (tempAddressWord.Right >= Left && tempAddressWord.Left <= Right)
                                        {
                                            //check Telphone number
                                            //if (isTelPhoneFound == false)
                                            //{
                                            if (Regex.IsMatch(tempAddressWord.strWord, rxTel.ToString()) == true && tempAddressWord.strWord.Trim().StartsWith("0") == false)
                                            {
                                                if (isTelPhoneFound == false)
                                                {
                                                    isTelPhoneFound = true;
                                                    row["TelPhone"] = tempAddressWord.strWord;
                                                    TempPhoneNo = tempAddressWord;
                                                    row["cncTelPhone"] = TempPhoneNo;
                                                    break;
                                                }
                                            }
                                            //}
                                            //************************

                                            if (iLoopCount.Contains(tempAddressWord.LineNo) == false)
                                            {
                                                iLoopCount.Add(tempAddressWord.LineNo);
                                            }


                                            //*******Added on 05March2021*********
                                            if (tempAddressWord.strWord.Trim().Split(' ').Length > 10)
                                            {
                                                nameLineNo = 0;
                                                break;
                                            }
                                            //*****************************************


                                            //*STEP5:  If SOLD Keyword found Ignore that address

                                            if (Regex.IsMatch(tempAddressWord.strWord, "sold", RegexOptions.IgnoreCase))
                                            {
                                                obj_clsCncWord_TempNameLinewise.Clear();
                                                obj_clsCncWord_TempName.Clear();
                                                obj_clsCncWord_TempAddress1.Clear();
                                                nameLineNo = 0;
                                                break;
                                            }
                                            else if (iWordNo > 1)
                                            {
                                                if (Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, "sold", RegexOptions.IgnoreCase) && tempAddressWord.Left - ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].Right < 300)
                                                // if (Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, "sold", RegexOptions.IgnoreCase))
                                                {
                                                    obj_clsCncWord_TempNameLinewise.Clear();
                                                    obj_clsCncWord_TempName.Clear();
                                                    obj_clsCncWord_TempAddress1.Clear();
                                                    nameLineNo = 0;
                                                    break;
                                                }
                                            }

                                            //STEP2: check Stop Keyword like Ship To / Consignee To / Bill To
                                            if (Regex.IsMatch(tempAddressWord.strWord, nameStopLinePattern, RegexOptions.IgnoreCase))
                                            {
                                                bNameStopLine = true;
                                                // nameLineNo = 0;
                                                //string strcncName = tempAddressWord.strWord;
                                                //strcncName = Module1.func_RemoveKeywordsforAllName(strcncName);
                                                //strcncName = Regex.Replace(strcncName.Trim(), "[^a-z0-9]", "",RegexOptions.IgnoreCase).Trim();

                                                //if (strcncName.Trim().Length < 3)
                                                //{
                                                //    continue;
                                                //}

                                                //if (obj_clsCncWord_TempName.Count > 0)
                                                //{
                                                //    nameLineNo = 0;
                                                //    break;
                                                //}

                                                if (obj_clsCncWord_TempName.Count > 0)
                                                {
                                                    if (Regex.IsMatch(tempAddressWord.strWord, "ship from|ship to|shipfrom|shipto", RegexOptions.IgnoreCase))
                                                    {
                                                        if (Regex.Replace(Regex.Replace(tempAddressWord.strWord, "ship from|ship to|shipfrom|shipto", "", RegexOptions.IgnoreCase), "[^a-z0-9]", "", RegexOptions.IgnoreCase).Trim().Length <= 2)
                                                        {
                                                            nameLineNo = 0;
                                                            break;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        nameLineNo = 0;
                                                        break;
                                                    }
                                                }

                                            }
                                            else if (iWordNo > 1)
                                            {
                                                if (Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, nameStopLinePattern, RegexOptions.IgnoreCase))
                                                {
                                                    bNameStopLine = true;
                                                    //if (obj_clsCncWord_TempName.Count > 0)
                                                    //{
                                                    // iWordNo = ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].WordCount + 1;
                                                    //nameLineNo = 0;
                                                    // break;
                                                    //}
                                                }
                                            }

                                            //*********************

                                            // STEP4:*****if bNameFound found still check on next line have Name line
                                            if (bNameFound == true && bNameStartWithNameKeyword_LineNo != tempAddressWord.LineNo)
                                            {
                                                if (tempAddressWord.strWord.ToLower().Contains("name") || tempAddressWord.strWord.ToLower().Contains("Narne"))
                                                {
                                                    string strnameLine = tempAddressWord.strWord;
                                                    strnameLine = Regex.Replace(strnameLine, "Narne:|Narne|name:|name", "", RegexOptions.IgnoreCase);
                                                    strnameLine = Regex.Replace(strnameLine, "[^a-z0-9]", "", RegexOptions.IgnoreCase);
                                                    if (strnameLine.Trim().Length >= 2)
                                                    {
                                                        obj_clsCncWord_TempNameLinewise.Add(tempAddressWord);
                                                    }
                                                    continue;
                                                }
                                                else if (iWordNo > 1)
                                                {
                                                    if (ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("name") || ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord.ToLower().Contains("Narne"))
                                                    {
                                                        string strnameLine = tempAddressWord.strWord;
                                                        strnameLine = Regex.Replace(strnameLine, "Narne:|Narne|name:|name", "", RegexOptions.IgnoreCase);
                                                        strnameLine = Regex.Replace(strnameLine, "[^a-z0-9]", "", RegexOptions.IgnoreCase);
                                                        if (strnameLine.Trim().Length >= 2)
                                                        {
                                                            obj_clsCncWord_TempNameLinewise.Add(tempAddressWord);
                                                        }
                                                        continue;
                                                    }
                                                    else
                                                    {
                                                        bNameStopLine = true;
                                                        nameLineNo = 0;
                                                        break;
                                                    }
                                                }
                                                else
                                                {
                                                    bNameStopLine = true;
                                                    nameLineNo = 0;
                                                    break;
                                                }
                                            }

                                            //******************************************************************************************

                                            //STEP3: check Name keyword Found 
                                            if (Regex.IsMatch(tempAddressWord.strWord, "Narne|name", RegexOptions.IgnoreCase))
                                            {
                                                bNameStartWithNameKeyword_LineNo = tempAddressWord.LineNo;
                                                bNameFound = true;
                                                if (Regex.Replace(tempAddressWord.strWord, "Narne:|Narne|name:|name", "", RegexOptions.IgnoreCase).Trim().Length < 3)
                                                {
                                                    continue;
                                                }
                                            }
                                            else if (iWordNo > 1)
                                            {

                                                if (Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, "name", RegexOptions.IgnoreCase) || Regex.IsMatch(ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].strWord, "Narne", RegexOptions.IgnoreCase))
                                                {
                                                    bNameStartWithNameKeyword_LineNo = tempAddressWord.LineNo;
                                                    bNameFound = true;
                                                }
                                            }

                                            //*********************************

                                            if (obj_clsCncWord_TempNameLinewise.Count > 0)
                                            {
                                                if (tempAddressWord.Left - obj_clsCncWord_TempNameLinewise[obj_clsCncWord_TempNameLinewise.Count - 1].Right > (obj_clsCncWord_TempNameLinewise[obj_clsCncWord_TempNameLinewise.Count - 1].AvgCharY1 * 4))
                                                {
                                                    break;
                                                }
                                            }

                                            //****** 18Jun21 if most left world right less than city left then world length >2
                                            if (tempAddressWord.Right < CityLeft && tempAddressWord.strWord.Trim().Length <= 2)
                                            {
                                                continue;
                                            }
                                            obj_clsCncWord_TempNameLinewise.Add(tempAddressWord);
                                        }
                                    }

                                    if (obj_clsCncWord_TempNameLinewise.Count > 0)
                                    {

                                        //*************17Jun21****************
                                        clsCnCWord objFinalWord = objCNC.MergeWords(obj_clsCncWord_TempNameLinewise.ToArray());

                                        if (obj_clsCncWord_TempName.Count > 0)// Name 2 line should not have more than 50 gap
                                        {
                                            if (objFinalWord.Left - obj_clsCncWord_TempName[obj_clsCncWord_TempName.Count - 1].Left > 50)
                                            {
                                                nameLineNo = 0;
                                                break;
                                            }
                                        }

                                        //************************************
                                        if (!regex_email.Match(objFinalWord.strWord).Success) //Line should not email id
                                        {
                                            if (!Module1.func_IgnoreKeyword(objFinalWord.strWord)) //Line should not email id
                                            {

                                                objFinalWord.strWord = Module1.func_RemoveKeywordsFromName(objFinalWord.strWord); //Remove Labels and Extream Left words from Word 
                                                                                                                                  //objFinalWord.strWord = Regex.Replace(objFinalWord.strWord.Trim(), " location#| location #| loc#| loc #", "", RegexOptions.IgnoreCase);

                                                if (Regex.Replace(objFinalWord.strWord, "[^a-zA-Z]", "").Trim().Length > 1) //Line Should contain atleast 2 alpha charactes 
                                                {
                                                    obj_clsCncWord_TempName.Add(objFinalWord);
                                                }
                                            }
                                        }
                                        //************************************

                                        // obj_clsCncWord_TempName.Add(objCNC.MergeWords(obj_clsCncWord_TempNameLinewise.ToArray()));
                                    }
                                    if (iLoopCount.Count > 3)
                                    {
                                        break;
                                    }
                                }


                                if (bNameFound == true)
                                {
                                    bNameStopLine = true;
                                }

                                if (bNameStopLine == false && obj_clsCncWord_TempName.Count > 1)
                                {
                                    string strFirstName = obj_clsCncWord_TempName[obj_clsCncWord_TempName.Count - 1].strWord.Trim();
                                    if (Regex.IsMatch(strFirstName, " LLC| INC| LTD|C/O", RegexOptions.IgnoreCase) == false)
                                    {
                                        obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);

                                        if (obj_clsCncWord_TempName.Count > 1)
                                        {
                                            strFirstName = obj_clsCncWord_TempName[obj_clsCncWord_TempName.Count - 1].strWord.Trim();
                                            if (Regex.IsMatch(strFirstName, " LLC| INC| LTD|C/O", RegexOptions.IgnoreCase) == false)
                                            {
                                                obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);

                                                if (obj_clsCncWord_TempName.Count > 1)
                                                {
                                                    strFirstName = obj_clsCncWord_TempName[obj_clsCncWord_TempName.Count - 1].strWord.Trim();
                                                    if (Regex.IsMatch(strFirstName, " LLC| INC| LTD|C/O", RegexOptions.IgnoreCase) == false)
                                                    {
                                                        obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);

                                                        if (obj_clsCncWord_TempName.Count > 1)
                                                        {
                                                            strFirstName = obj_clsCncWord_TempName[obj_clsCncWord_TempName.Count - 1].strWord.Trim();
                                                            if (Regex.IsMatch(strFirstName, " LLC| INC| LTD|C/O", RegexOptions.IgnoreCase) == false)
                                                            {
                                                                obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    //obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);
                                    //if (obj_clsCncWord_TempName.Count > 1)
                                    //{
                                    //    obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);
                                    //}
                                    //if (obj_clsCncWord_TempName.Count > 1)
                                    //{
                                    //    obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);
                                    //}
                                    //if (obj_clsCncWord_TempName.Count > 1)
                                    //{
                                    //    obj_clsCncWord_TempName.RemoveAt(obj_clsCncWord_TempName.Count - 1);
                                    //}
                                }

                            }
                        }
                        #endregion

                        if (obj_clsCncWord_TempAddress1.Count == 0)
                        {
                            continue;
                        }

                        obj_clsCncWord_TempAddress1.Reverse();
                        obj_clsCncWord_TempName.Reverse();

                        #region"AddressType"

                        if (obj_clsCncWord_TempAddress1.Count > 0)
                        {

                            int AddressLineNo = obj_clsCncWord_TempAddress1[0].LineNo;

                            if (AddressLineNo > 1)
                            {
                                //string nameStopLinePattern = "Name|From| To|Location|Shipping|shipper|Consigned|Cons||Deliver|Destination|PPD|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight";



                                string strShipFrom = "From| At |shipper|Origin|Pick up|Pickup";//Shipping
                                                                                               //string strShipTo = "Consign|Ship To|Deliver|Destination|Shipped To|Sold To";
                                string strShipTo = "Consign|ShipTo|Ship To|Deliver|Destination|ShippedTo|Shipped To|Receiver";
                                // string strBillTo = "Payment To|Charges To|PPD|Bill To|SFBT|Payor|Invoice|Third|3rd|Party|Remit|Freight";
                                string strBillTo = "Payment To|Charges To|PPD|BillTo|Bill To|SFBT|Third|3rd|Party|Remit|freight to|REMITTO|REMIT TO|THIRD PARTY|FREIGHT CHARGES|BILLED TO|PREPAID CHARGE|B/L to";//new

                                int Left = obj_clsCncWord_TempAddress1[0].Left - 100;
                                if (Left < 0)
                                {
                                    Left = 0;
                                }
                                int Right = obj_clsCncWord_TempAddress1[0].Right + 300;
                                //int iLoopCount = 0;
                                ArrayList iLoopCount = new ArrayList();
                                int y2Top = obj_clsCncWord_TempAddress1[0].Top;

                                ObjMetaData2 = iQPUBLIC.PublicComponents.oMetaData2;

                                for (int nameLineNo = AddressLineNo; 0 < nameLineNo; nameLineNo--)
                                {

                                    if (iLoopCount.Count > 3)
                                    {
                                        break;
                                    }

                                    for (int iWordNo = 1; iWordNo <= ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].WordCount; iWordNo++)
                                    {

                                        clsCnCWord objtempWord = (clsCnCWord)ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo].DeepClone();

                                        //STEP1: AddressLine Within CityStateZip RIO
                                        if (objtempWord.Right >= Left && objtempWord.Left <= Right && objtempWord.strWord.Split(' ').Length < 10 && y2Top - objtempWord.Bottom < 300)
                                        {

                                            //iLoopCount += 1;

                                            if (iLoopCount.Contains(objtempWord.LineNo) == false)
                                            {
                                                iLoopCount.Add(objtempWord.LineNo);
                                            }

                                            //else
                                            //{


                                            if (Regex.IsMatch(objtempWord.strWord, strShipFrom, RegexOptions.IgnoreCase) == true)
                                            {
                                                TempAddressType = objtempWord;
                                                TempAddressType.strWord = "SHIPPER";
                                                row["AddressType"] = "SHIPPER";
                                                row["cncAddressType"] = TempAddressType;
                                                nameLineNo = 0;
                                                break;
                                            }
                                            else if (Regex.IsMatch(objtempWord.strWord, strShipTo, RegexOptions.IgnoreCase) == true)
                                            {
                                                TempAddressType = objtempWord;
                                                TempAddressType.strWord = "CONSIGNEE";
                                                row["AddressType"] = "CONSIGNEE";
                                                row["cncAddressType"] = TempAddressType;
                                                nameLineNo = 0;
                                                if ((Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.75)) > objtempWord.Top)
                                                {
                                                    ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T"; // Added on 6-July-2021 => Tag consignee only when if we found keyword or matched with Pick up DB
                                                }
                                                break;
                                            }
                                            else if (Regex.IsMatch(objtempWord.strWord, strBillTo, RegexOptions.IgnoreCase) == true)
                                            {

                                                TempAddressType = objtempWord;
                                                TempAddressType.strWord = "3RD PARTY";
                                                row["AddressType"] = "3RD PARTY";
                                                row["cncAddressType"] = TempAddressType;
                                                nameLineNo = 0;
                                                break;
                                            }

                                            else if (iWordNo > 1)
                                            {
                                                clsCnCWord objtempWord2 = (clsCnCWord)ObjMetaData2.Page[intCurrPageNumber].Line[nameLineNo].Word[iWordNo - 1].DeepClone();

                                                //string DataLabel = objtempWord2.strWord.Trim().Split(' ')[objtempWord2.strWord.Trim().Split(' ').Length - 1];
                                                if (objtempWord2.strWord.Split(' ').Length < 10 && objtempWord.Left - objtempWord2.Left <= 500)
                                                {
                                                    string DataLabel = objtempWord2.strWord;
                                                    if (Regex.IsMatch(DataLabel, strShipFrom, RegexOptions.IgnoreCase) == true)
                                                    {
                                                        TempAddressType = objtempWord2;
                                                        TempAddressType.strWord = "SHIPPER";
                                                        row["AddressType"] = "SHIPPER";
                                                        row["cncAddressType"] = TempAddressType;
                                                        nameLineNo = 0;
                                                        break;
                                                    }
                                                    else if (Regex.IsMatch(DataLabel, strShipTo, RegexOptions.IgnoreCase) == true)
                                                    {
                                                        TempAddressType = objtempWord2;
                                                        TempAddressType.strWord = "CONSIGNEE";
                                                        row["AddressType"] = "CONSIGNEE";
                                                        row["cncAddressType"] = TempAddressType;
                                                        nameLineNo = 0;
                                                        if ((Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.75)) > objtempWord2.Top)
                                                        {
                                                            ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T"; // Added on 6-July-2021 => Tag consignee only when if we found keyword or matched with Pick up DB
                                                        }
                                                        break;
                                                    }
                                                    else if (Regex.IsMatch(DataLabel, strBillTo, RegexOptions.IgnoreCase) == true)
                                                    {

                                                        TempAddressType = objtempWord2;
                                                        TempAddressType.strWord = "3RD PARTY";
                                                        row["AddressType"] = "3RD PARTY";
                                                        row["cncAddressType"] = TempAddressType;
                                                        nameLineNo = 0;
                                                        break;
                                                    }
                                                }
                                            }

                                            // }
                                        }
                                    }
                                }
                                //


                            }
                        }
                        #endregion

                        #region "Check Hight from Name if LineCount>1"

                        if (obj_clsCncWord_TempName.Count > 1)
                        {
                            int maxHight = 0;

                            maxHight = TempCityValue.Top - obj_clsCncWord_TempAddress1[obj_clsCncWord_TempAddress1.Count - 1].Bottom;

                            if (obj_clsCncWord_TempAddress1.Count > 1)
                            {
                                //for (int addLineNo = 1; addLineNo <= obj_clsCncWord_TempAddress1.Count - 1; addLineNo++)
                                for (int addLineNo = obj_clsCncWord_TempAddress1.Count - 2; 0 <= addLineNo; addLineNo--)
                                {
                                    if (obj_clsCncWord_TempAddress1[addLineNo + 1].Top - obj_clsCncWord_TempAddress1[addLineNo].Bottom > maxHight)
                                    {
                                        maxHight = obj_clsCncWord_TempAddress1[addLineNo + 1].Top - obj_clsCncWord_TempAddress1[addLineNo].Bottom;
                                    }
                                }
                            }


                            int stopLine = -1;
                            for (int addLineNo = obj_clsCncWord_TempName.Count - 2; 0 <= addLineNo; addLineNo--)
                            {
                                if (obj_clsCncWord_TempName[addLineNo + 1].Top - obj_clsCncWord_TempName[addLineNo].Bottom > maxHight + 10)
                                {
                                    stopLine = addLineNo;
                                    break;
                                }
                            }

                            if (stopLine >= 0)
                            {
                                if (stopLine == 0)
                                {
                                    obj_clsCncWord_TempName.RemoveAt(0);
                                }
                                else if (stopLine == 1)
                                {
                                    obj_clsCncWord_TempName.RemoveAt(0);
                                    obj_clsCncWord_TempName.RemoveAt(0);

                                }
                            }


                        }
                        #endregion


                        #region "remove telphone number if it above on name"
                        if (isTelPhoneFound == true && obj_clsCncWord_TempName.Count > 0)
                        {
                            if (obj_clsCncWord_TempName[0].LineNo > TempPhoneNo.LineNo)
                            {
                                isTelPhoneFound = false;
                                row["TelPhone"] = "";
                                TempPhoneNo = null;
                                row["cncTelPhone"] = TempPhoneNo;
                            }
                        }
                        #endregion 

                        //**************Added on 05March21****
                        if (objCNC.MergeWords(obj_clsCncWord_TempAddress1.ToArray()).strWord.Trim().Length <= 5)
                        {
                            continue;
                        }
                        //**************************************

                        // Temp code commented and added below on 28-Nov-2020 -- start
                        //row["AddressLine1"] = StrBuilderaddressLine1;
                        //row["Name"] = StrBuilderName;
                        if (obj_clsCncWord_TempAddress1.Count > 0)
                        {
                            row["AddressLine1"] = objCNC.MergeWords(obj_clsCncWord_TempAddress1.ToArray()).strWord;
                        }

                        //if (obj_clsCncWord_TempName.Count > 0)
                        //{
                        //    row["Name"] = objCNC.MergeWords(obj_clsCncWord_TempName.ToArray()).strWord;
                        //}
                        // Temp code commented and added below on 28-Nov-2020 -- end 

                        if (obj_clsCncWord_TempAddress1.Count > 0)
                        {
                            obj_clsCncWord_Temp.AddRange(obj_clsCncWord_TempAddress1);
                        }

                        obj_clsCncWord_Temp.Add(TempCityValue);
                        obj_clsCncWord_Temp.Add(TempStateValue);
                        //obj_clsCncWord_Temp.Add(TempZipValue);//commanted by RajeshB on 28DEC20

                        //***********RajeshB Added 28DEC20
                        //if (TempStateValue.strWord.Trim().Length <= 4)
                        //{
                        obj_clsCncWord_Temp.Add(TempZipValue);
                        //}
                        //********************************

                        row["MergeWords"] = objCNC.MergeWords(obj_clsCncWord_Temp.ToArray()).strWord;
                        row["clcncMergeWords"] = objCNC.MergeWords(obj_clsCncWord_Temp.ToArray());

                        List<clsCnCWord> objCityStateZip = new List<clsCnCWord>();
                        objCityStateZip.Add(TempCityValue);
                        objCityStateZip.Add(TempStateValue);
                        //objCityStateZip.Add(TempZipValue);//commanted by RajeshB on 28DEC20

                        //***********RajeshB Added 28DEC20
                        //if (TempStateValue.strWord.Trim().Length <= 4)
                        //{
                        objCityStateZip.Add(TempZipValue);
                        //}

                        //if (obj_clsCncWord_TempName.Count > 0)
                        //{
                        //    row["cncName"] = objCNC.MergeWords(obj_clsCncWord_TempName.ToArray());
                        //}
                        //if (obj_clsCncWord_TempName.Count > 0)
                        //{
                        //    row["Name"] = objCNC.MergeWords(obj_clsCncWord_TempName.ToArray()).strWord;
                        //}
                        if (obj_clsCncWord_TempName.Count == 0)
                        {
                            row["cncName"] = GetDefaultWord(intCurrPageNumber, "*****");
                            row["Name"] = "*****";
                        } // above if part is added on 7-jan-2021 
                        else if (obj_clsCncWord_TempName.Count > 0)
                        {
                            clsCnCWord cncName = objCNC.MergeWords(obj_clsCncWord_TempName.ToArray());
                            string strcncName = cncName.strWord;
                            if (strcncName.Split(' ').Length > 10)
                            {
                                // MessageBox.Show("NameLenght>10: " + cncName.strWord);
                                //continue;
                                row["cncName"] = GetDefaultWord(intCurrPageNumber, "*****");
                                row["Name"] = "*****";
                            }
                            else
                            {
                                //*********************If name have special char and numeric then return ****** only
                                //string RemoveWordsName = "3RD PARTY BILL|3RD PARTY BILL FREIGHT PREPAID|BILL OR REMIT|REMIT|BILL FREIGHT CHARGES|BILL THIRD PARTY INFORMATION|FREIGHT PAYMELT PLAN|FREIGHT CHARGE TERMS (FREIGHT CHARGES ARE PREPAID|FREIGHT CHARGES BILL";
                                strcncName = Module1.func_RemoveKeywordsforAllName(strcncName);

                                strcncName = Regex.Replace(strcncName.Trim(), "^[^a-zA-Z0-9]+", "").Trim();
                                //Regex rxStreetNumber = new Regex("^[0-9]+$");
                                int TotCharCount = strcncName.Replace(" ", "").Length;
                                int TotAlphaNumericCount = strcncName.Replace(" ", "").Count(char.IsLetterOrDigit);
                                int TotRemainingSpecialCharCount = TotCharCount - TotAlphaNumericCount;
                                int TotNumericCount = strcncName.Replace(" ", "").Count(char.IsDigit);
                                // Check name len should not be zero or it should be greater than 2 alphabet
                                string abc = Regex.Replace(strcncName, "[^a-z0-9\\s]", "", RegexOptions.IgnoreCase).Trim();
                                string tempname = " " + abc.Split(' ')[abc.Split(' ').Length - 1] + " ";
                                if (Regex.Replace(strcncName, "[^a-z]", "", RegexOptions.IgnoreCase).Trim().Length == 0 || Regex.Replace(strcncName, "[^a-z]", "", RegexOptions.IgnoreCase).Trim().Length <= 2)
                                {
                                    row["cncName"] = GetDefaultWord(intCurrPageNumber, "*****");
                                    row["Name"] = "*****";
                                }
                                // if 5 or more than 5 special characters are present then don't consider as valid name 
                                else if (TotRemainingSpecialCharCount >= 10)
                                {
                                    row["cncName"] = GetDefaultWord(intCurrPageNumber, "*****");
                                    row["Name"] = "*****";
                                }

                                // if name contains BillTo/Remit/ 3rd party bill then don't consider as valid name
                                //else if (Regex.IsMatch(Module1.func_RemoveKeywordsforAllName(strcncName), RemoveWordsName, RegexOptions.IgnoreCase) == true)
                                //{
                                //    row["cncName"] = GetDefaultWord(intCurrPageNumber, "*****");
                                //    row["Name"] = "*****";
                                //}
                                // if name start with numeric OR ends with street then don't consider as valid name, it might be address starts with street number and 
                                //else if (!string.IsNullOrEmpty(strcncName) && strcncName.Split(' ')[0].Trim().Length > 0 && (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(strcncName.Split(' ')[0]).Trim(), rxStreetNumber.ToString()) == true || Regex.IsMatch(Module1.func_RemoveSpecialCharacter(strcncName.Split(' ')[strcncName.Split(' ').Length - 1]).Trim(), AbbrivationPattern_Endwith) == true))

                                //else if (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(strcncName.Trim().Split(' ')[0]).Trim(), rxStreetNumber.ToString()) == true || Regex.IsMatch(tempname, AbbrivationPattern) == true)

                                //*******************AbbrivationPattern*********************
                                //else if (Regex.IsMatch(tempname, AbbrivationPattern) == true)
                                //{
                                //    row["cncName"] = GetDefaultWord(intCurrPageNumber, "*****");
                                //    row["Name"] = "*****";
                                //}
                                //****************AbbrivationPattern****************************
                                // check total numeric count is greater than 3 then dont consider as valid name
                                //else if (TotNumericCount > 3)
                                //{
                                //    row["cncName"] = GetDefaultWord(intCurrPageNumber, "*****");
                                //    row["Name"] = "*****";
                                //}
                                else
                                {
                                    row["cncName"] = cncName;
                                    row["Name"] = cncName.strWord;
                                }

                            }
                        }

                        //if (obj_clsCncWord_TempAddress1.Count > 0)
                        //{
                        //    row["cncAddress"] = objCNC.MergeWords(obj_clsCncWord_TempAddress1.ToArray());
                        //    row["LineNo"] = obj_clsCncWord_TempAddress1[0].LineNo;
                        //}
                        if (obj_clsCncWord_TempAddress1.Count == 0)
                        {
                            continue;
                        }
                        else if (obj_clsCncWord_TempAddress1.Count > 0)
                        {

                            clsCnCWord cncAddr = objCNC.MergeWords(obj_clsCncWord_TempAddress1.ToArray());
                            if (cncAddr.strWord.Split(' ').Length > 25 || cncAddr.strWord.Trim() == "")
                            {
                                // MessageBox.Show("AddreLenght>20: " + cncAddr.strWord);
                                continue;
                            }
                            row["cncAddress"] = cncAddr;
                            row["LineNo"] = obj_clsCncWord_TempAddress1[0].LineNo;
                        }

                        if (objCityStateZip.Count > 0)
                        {
                            row["cncCityStateZip"] = objCNC.MergeWords(objCityStateZip.ToArray());
                        }


                        Module1.TableRecordss.Rows.Add(row);

                    }



                }
            }

            Module1.TableRecordss.DefaultView.Sort = "LineNo";
            Module1.TableRecordss = Module1.TableRecordss.DefaultView.ToTable();
            #endregion

            #region Commented on 30-July-2021
            //if (Module1.TableRecordss.Rows.Count > 0)
            //{
            //    Regex rxOnlyAlphaNumeric = new Regex("^(?=.*[0-9])(?=.*[a-zA-Z])([a-zA-Z0-9]+)$"); // Only alphanumeric
            //    Regex rxNumeric = new Regex("^[0-9]+$");
            //    Regex rxAlphabet = new Regex("^[a-zA-Z]+$");
            //    string State = String.Empty;
            //    string ZipCode = String.Empty;
            //    string City = string.Empty;
            //    string AddressLine1 = string.Empty;
            //    string AddressLine2 = string.Empty;
            //    string Name = string.Empty;
            //    string AddressType = "";
            //    string MergeWords = string.Empty;
            //    int LineNo = 0;

            //    //clsCnCWord clcncMergeWords = new clsCnCWord();
            //    //clsCnCWord cncName = new clsCnCWord();
            //    //clsCnCWord cncAddress = new clsCnCWord();
            //    //clsCnCWord cncCityStateZip = new clsCnCWord();
            //    //clsCnCWord cncTelPhone = new clsCnCWord();
            //    //clsCnCWord cncAddressType = new clsCnCWord();
            //    // Check shipper details
            //    string ShipperDetailsAddressLine = string.Empty;
            //    string ShipperDetailsCity = string.Empty;
            //    string ShipperDetailsState = string.Empty;
            //    string ShipperDetailsZip = string.Empty;
            //    string ConsigneeDetailsAddressLine = string.Empty;
            //    string ConsigneeDetailsCity = string.Empty;
            //    string ConsigneeDetailsState = string.Empty;
            //    string ConsigneeDetailsZip = string.Empty;
            //    string BillToDetailsAddressLine = string.Empty;
            //    string BillToDetailsCity = string.Empty;
            //    string BillToDetailsState = string.Empty;
            //    string BillToDetailsZip = string.Empty;
            //    // Added on 23-July-2021 => To fix wrong tagging issue
            //    string ShipperDetailsName = string.Empty;
            //    string ConsigneeDetailsName = string.Empty;
            //    string BillToDetailsName = string.Empty;
            //    if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSET2_PickupShipperDetails"))
            //    {
            //        DataTable dtKeys = new DataTable();
            //        dtKeys = (DataTable)iQPUBLIC.PublicComponents.htMyVariable["BOLSET2_PickupShipperDetails"];
            //        if (dtKeys != null && dtKeys.Rows.Count > 0)
            //        {
            //            ShipperDetailsAddressLine = dtKeys.Rows[0]["PASA1"].ToString().Trim() + dtKeys.Rows[0]["PASA2"].ToString().Trim();
            //            ShipperDetailsCity = dtKeys.Rows[0]["PASCT"].ToString().Trim();
            //            ShipperDetailsCity = ShipperDetailsCity.Replace("'", ""); // Added on 10-March-2020 => To handle exception in the query when city with an apostrophe in it	
            //            ShipperDetailsZip = dtKeys.Rows[0]["PASZIP"].ToString().Trim();
            //            ShipperDetailsName = dtKeys.Rows[0]["PASNM"].ToString().Trim(); // Added on 23-July-2021 => To fix wrong tagging issue
            //            // Consignee
            //            ConsigneeDetailsAddressLine = dtKeys.Rows[0]["ConsigneeAddress1"].ToString().Trim() + dtKeys.Rows[0]["ConsigneeAddress2"].ToString().Trim();
            //            ConsigneeDetailsCity = dtKeys.Rows[0]["ConsigneeCity"].ToString().Trim();
            //            ConsigneeDetailsCity = ConsigneeDetailsCity.Replace("'", ""); // Added on 10-March-2020 => To handle exception in the query when city with an apostrophe in it
            //            ConsigneeDetailsZip = dtKeys.Rows[0]["ConsigneeZip"].ToString().Trim();
            //            ConsigneeDetailsName = dtKeys.Rows[0]["ConsigneeName"].ToString().Trim(); // Added on 23-July-2021 => To fix wrong tagging issue
            //            // BillTo
            //            BillToDetailsAddressLine = dtKeys.Rows[0]["BillToAddress1"].ToString().Trim() + dtKeys.Rows[0]["BillToAddress2"].ToString().Trim();
            //            BillToDetailsCity = dtKeys.Rows[0]["BillToCity"].ToString().Trim();
            //            BillToDetailsCity = BillToDetailsCity.Replace("'", ""); // Added on 10-March-2020 => To handle exception in the query when city with an apostrophe in it
            //            BillToDetailsZip = dtKeys.Rows[0]["BillToZipCode"].ToString().Trim();
            //            BillToDetailsName = dtKeys.Rows[0]["BillToName"].ToString().Trim(); // Added on 23-July-2021 => To fix wrong tagging issue
            //        }
            //    }
            //    //bool isShipperDouble = false;
            //    //bool isConsigneeDouble = false;
            //    //bool isThirdPartyDouble = false;
            //    //DataRow[] RowsShipper = Module1.TableRecordss.Select("[AddressType] = 'SHIPPER'");
            //    //if (RowsShipper.Length > 1)
            //    //{
            //    //    isShipperDouble = true;
            //    //}
            //    //DataRow[] RowsConsignee = Module1.TableRecordss.Select("[AddressType] = 'CONSIGNEE'");
            //    //if (RowsConsignee.Length > 1)
            //    //{
            //    //    isConsigneeDouble = true;
            //    //}
            //    //DataRow[] RowsThirdParty = Module1.TableRecordss.Select("[AddressType] = '3RD PARTY'");
            //    //if (RowsThirdParty.Length > 1)
            //    //{
            //    //    isThirdPartyDouble = true;
            //    //}
            //    bool isShipperTag = false;
            //    bool isConsigneeTag = false;
            //    bool isBillToTag = false;

            //    //bool currShipperRow = false;
            //    //bool currConsigneeRow = false;
            //    //bool currThirdPartyRow = false;
            //    for (int i = 0; i < Module1.TableRecordss.Rows.Count; i++)
            //    {
            //        bool isValidZip = false;
            //        bool isValidCity = false;
            //        bool isNOTValidState = false;
            //        bool isValidAddressLine1 = false;
            //        clsCnCWord clcncMergeWords = new clsCnCWord();
            //        clsCnCWord cncName = new clsCnCWord();
            //        clsCnCWord cncAddress = new clsCnCWord();
            //        clsCnCWord cncCity = new clsCnCWord();
            //        clsCnCWord cncState = new clsCnCWord();
            //        clsCnCWord cncZip = new clsCnCWord();
            //        clsCnCWord cncCityStateZip = new clsCnCWord();
            //        clsCnCWord cncTelPhone = new clsCnCWord();
            //        clsCnCWord cncAddressType = new clsCnCWord();
            //        DataRow rowFinal = Module1.TableRecordssFinal.NewRow();

            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["Sate"]))
            //        {
            //            State = Module1.TableRecordss.Rows[i]["Sate"].ToString();
            //        }
            //        else
            //        {
            //            State = "";
            //        }
            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["Zip"]))
            //        {
            //            ZipCode = Module1.TableRecordss.Rows[i]["Zip"].ToString();
            //        }
            //        else
            //        {
            //            ZipCode = "";
            //        }

            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["City"]))
            //        {
            //            City = Module1.TableRecordss.Rows[i]["City"].ToString();
            //        }
            //        else
            //        {
            //            City = "";
            //        }
            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["AddressLine1"]))
            //        {
            //            AddressLine1 = Module1.TableRecordss.Rows[i]["AddressLine1"].ToString();
            //        }
            //        else
            //        {
            //            AddressLine1 = "";
            //        }
            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["AddressLine2"]))
            //        {
            //            AddressLine2 = Module1.TableRecordss.Rows[i]["AddressLine2"].ToString();
            //        }
            //        else
            //        {
            //            AddressLine2 = "";
            //        }
            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["Name"]))
            //        {
            //            Name = Module1.TableRecordss.Rows[i]["Name"].ToString();
            //        }
            //        else
            //        {
            //            Name = "";
            //        }

            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["AddressType"]))
            //        {
            //            AddressType = Module1.TableRecordss.Rows[i]["AddressType"].ToString();

            //            cncAddressType = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncAddressType"];

            //        }
            //        else
            //        {
            //            AddressType = "";
            //            //cncAddressType = ;
            //        }

            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["MergeWords"]))
            //        {
            //            MergeWords = Module1.TableRecordss.Rows[i]["MergeWords"].ToString();
            //            clcncMergeWords = (clsCnCWord)Module1.TableRecordss.Rows[i]["clcncMergeWords"];

            //            cncCityStateZip = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncCityStateZip"];
            //        }
            //        else
            //        {
            //            MergeWords = "";
            //        }

            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncName"]))
            //        {
            //            cncName = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncName"];
            //        }

            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncAddress"]))
            //        {
            //            cncAddress = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncAddress"];
            //        }
            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncCity"]))
            //        {
            //            cncCity = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncCity"];
            //        }
            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncState"]))
            //        {
            //            cncState = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncState"];
            //        }
            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncZip"]))
            //        {
            //            cncZip = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncZip"];
            //        }

            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncTelPhone"]))
            //        {
            //            cncTelPhone = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncTelPhone"];
            //        }
            //        if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["LineNo"]))
            //        {
            //            LineNo = Convert.ToInt32(Module1.TableRecordss.Rows[i]["LineNo"]);
            //        }
            //        else
            //        {
            //            LineNo = 0;
            //        }


            //        if (!string.IsNullOrEmpty(State))
            //        {
            //            if (State == "ID:")
            //            {
            //                isNOTValidState = true;
            //            }
            //        }


            //        if (!string.IsNullOrEmpty(ZipCode))
            //        {
            //            // check the valid zip code eg:- 43125-1301  then pick the first part of the zip
            //            // below code commented on 21-Jan-2021 -- start
            //            //if (ZipCode.Contains('-') == true)
            //            //{
            //            //    string[] ZipValue = ZipCode.Split('-');
            //            //    if (ZipValue[0].Length <= 5)
            //            //    {
            //            //        ZipCode = ZipValue[0];
            //            //    }
            //            //}
            //            // below code commented on 21-Jan-2021 -- end
            //            // Code Added on 21-Jan-2021 -- start
            //            if (Module1.func_RemoveSpecialCharacter(ZipCode).Length >= 5)
            //            {
            //                ZipCode = ZipCode.Substring(0, 5);
            //            }
            //            // Code Added on 21-Jan-2021 -- end
            //            if ((Regex.IsMatch(Module1.func_RemoveSpecialCharacter(ZipCode).Trim(), rxNumeric.ToString()) == true || Regex.IsMatch(Module1.func_RemoveSpecialCharacter(ZipCode).Trim(), rxOnlyAlphaNumeric.ToString()) == true) && Module1.func_RemoveSpecialCharacter(ZipCode).Length >= 3)
            //            {

            //                isValidZip = true;

            //                if (ZipCode.Contains('.') == true)
            //                {
            //                    isValidZip = false;
            //                }
            //            }
            //        }
            //        if (!string.IsNullOrEmpty(City))
            //        {
            //            if ((Regex.IsMatch(Module1.func_RemoveSpecialCharacter(City).Trim(), rxAlphabet.ToString()) == true || Regex.IsMatch(Module1.func_RemoveSpecialCharacter(City).Trim(), rxOnlyAlphaNumeric.ToString()) == true) && Module1.func_RemoveSpecialCharacter(City).Length >= 3)
            //            {
            //                isValidCity = true;
            //            }
            //        }
            //        if (!string.IsNullOrEmpty(AddressLine1))
            //        {
            //            // check the abbrivation in the AddressLine1 and check numeric with two capital Alphabet
            //            if (Abbrivation.Any(AddressLine1.ToUpper().Contains))
            //            {
            //                isValidAddressLine1 = true;
            //            }
            //        }

            //        if ((isValidZip && isValidCity && !isNOTValidState))
            //        {
            //            rowFinal["Sate"] = State;
            //            rowFinal["Zip"] = ZipCode;
            //            rowFinal["City"] = City;

            //            if (AddressLine1 != null)
            //            {
            //                AddressLine1 = Module1.func_RemoveKeywordsNameAddress(AddressLine1);
            //            }
            //            rowFinal["AddressLine1"] = AddressLine1;
            //            rowFinal["AddressLine2"] = AddressLine2;

            //            if (Name != null)
            //            {
            //                //Name = Module1.func_RemoveKeywordsNameAddress(Name);
            //                Name = Module1.func_RemoveKeywordsforAllName(Name);
            //                //**********08FEB21
            //                Name = Regex.Replace(Name, "^[^a-zA-Z0-9%]+", "");
            //                Name = Regex.Replace(Name, "[^a-zA-Z0-9)]+$", "");
            //            }
            //            rowFinal["Name"] = Name;

            //            rowFinal["MergeWords"] = MergeWords;
            //            rowFinal["clcncMergeWords"] = clcncMergeWords;
            //            if (cncName != null)
            //            {
            //                if (cncName.strWord != null)
            //                {
            //                    // cncName.strWord = Module1.func_RemoveKeywordsNameAddress(cncName.strWord);
            //                    cncName.strWord = Module1.func_RemoveKeywordsforAllName(cncName.strWord);

            //                    //**********08FEB21
            //                    cncName.strWord = Regex.Replace(cncName.strWord, "^[^a-zA-Z0-9%]+", "");
            //                    cncName.strWord = Regex.Replace(cncName.strWord, "[^a-zA-Z0-9)]+$", "");
            //                }
            //            }
            //            rowFinal["cncName"] = cncName;
            //            rowFinal["cncCity"] = cncCity;
            //            rowFinal["cncState"] = cncState;
            //            if (cncZip != null)
            //            {
            //                if (cncZip.strWord != null)
            //                {
            //                    // check the valid zip code eg:- 43125-1301  then pick the first part of the zip
            //                    // below code is commnted on 21-Jan-2021 -- start
            //                    //if (cncZip.strWord.Contains('-') == true)
            //                    //{
            //                    //    string[] ZipValue = ZipCode.Split('-');
            //                    //    if (ZipValue[0].Length <= 5)
            //                    //    {
            //                    //        cncZip.strWord = ZipValue[0];
            //                    //    }
            //                    //}
            //                    // below code is commnted on 21-Jan-2021 -- end
            //                    // Code Added on 21-Jan-2021 -- start
            //                    if (Module1.func_RemoveSpecialCharacter(cncZip.strWord).Length >= 5)
            //                    {
            //                        cncZip.strWord = cncZip.strWord.Substring(0, 5);
            //                    }
            //                    // Code Added on 21-Jan-2021 -- end
            //                }
            //            }
            //            rowFinal["cncZip"] = cncZip;
            //            if (cncAddress != null)
            //            {
            //                if (cncAddress.strWord != null)
            //                {
            //                    cncAddress.strWord = Module1.func_RemoveKeywordsNameAddress(cncAddress.strWord);

            //                    //**********08FEB21
            //                    cncAddress.strWord = Regex.Replace(cncAddress.strWord, "^[^a-zA-Z0-9%]+", "");
            //                    cncAddress.strWord = Regex.Replace(cncAddress.strWord, "[^a-zA-Z0-9)]+$", "");
            //                }
            //            }
            //            rowFinal["cncAddress"] = cncAddress;
            //            rowFinal["cncCityStateZip"] = cncCityStateZip;
            //            if (cncTelPhone != null)
            //            {
            //                if (cncTelPhone.strWord != null)
            //                {
            //                    string str_DataTel = cncTelPhone.strWord;
            //                    string Regxtel = "[^0-9]";
            //                    cncTelPhone.strWord = Regex.Replace(str_DataTel, Regxtel, "", RegexOptions.IgnoreCase);
            //                }
            //            }
            //            rowFinal["cncTelPhone"] = cncTelPhone;
            //            rowFinal["cncAddressType"] = cncAddressType;

            //            // rowFinal["LineNo"] = LineNo;

            //            //if (AddressType == "SHIPPER")
            //            //{
            //            //    currShipperRow = true;
            //            //}
            //            //else if (AddressType == "CONSIGNEE")
            //            //{
            //            //    currConsigneeRow = true;
            //            //}
            //            //else if (AddressType == "3RD PARTY")
            //            //{
            //            //    currThirdPartyRow = true;
            //            //}

            //            if (clsF27.clsModule.clsPickupValuesFRP001RAT != "Y")
            //            {
            //                // If AddressType is not tag as shipper,consignee or bill to then apply ping pong with the captured data
            //                #region Ping-pong using pick up => FRP001
            //                //if (AddressType == "" || isShipperDouble == true || isConsigneeDouble == true || isThirdPartyDouble == true)
            //                //{
            //                bool isAddressMatch = false;
            //                bool isCityMatch = false;
            //                bool isZipCodeMatch = false;

            //                bool isAddressMatchConsignee = false;
            //                bool isCityMatchConsignee = false;
            //                bool isZipCodeMatchConsignee = false;

            //                bool isAddressMatchBillTo = false;
            //                bool isCityMatchBillTo = false;
            //                bool isZipCodeMatchBillTo = false;

            //                // Added on 23-July-2021 => To avoid wrong tagging
            //                bool isNameMatch = false;
            //                bool isNameMatchConsignee = false;
            //                bool isNameMatchBillTo = false;
            //                if (isShipperTag == false)
            //                {
            //                    if (!string.IsNullOrEmpty(ShipperDetailsAddressLine))
            //                    {
            //                        //objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ShipperDetailsAddressLine), Module1.func_RemoveSpecialCharacter(AddressLine1));
            //                        //if (objRet26.PercentageMatch >= 50)
            //                        //{
            //                        //    isAddressMatch = true;
            //                        //}

            //                        AddressLine1 = clsSearchCustomer.ReplaceAddressAbbreviation(AddressLine1.Trim(), Module1.dtAbbreviationTable);
            //                        ShipperDetailsAddressLine = clsSearchCustomer.ReplaceAddressAbbreviation(ShipperDetailsAddressLine.Trim(), Module1.dtAbbreviationTable);

            //                        var arrShipperDetailsAddressLine = ShipperDetailsAddressLine.Split(' ');
            //                        var arrAddressLine1 = AddressLine1.Split(' ');
            //                        int shipperMatchedCount = arrShipperDetailsAddressLine.Count(x => arrAddressLine1.Contains(x));
            //                        if (shipperMatchedCount >= 1 && arrShipperDetailsAddressLine.Length >= 1)
            //                        {
            //                            isAddressMatch = true;
            //                        }
            //                    }

            //                    // City 
            //                    if (!string.IsNullOrEmpty(ShipperDetailsCity))
            //                    {
            //                        //objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ShipperDetailsCity), Module1.func_RemoveSpecialCharacter(City));
            //                        objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(City), Module1.func_RemoveSpecialCharacter(ShipperDetailsCity));
            //                        if (objRet26.PercentageMatch >= 80)
            //                        {
            //                            isCityMatch = true;
            //                        }
            //                    }
            //                    // Zip 
            //                    if (!string.IsNullOrEmpty(ShipperDetailsZip))
            //                    {
            //                        objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ShipperDetailsZip), Module1.func_RemoveSpecialCharacter(ZipCode));
            //                        if (objRet26.PercentageMatch >= 80)
            //                        {
            //                            isZipCodeMatch = true;
            //                        }
            //                    }
            //                    // Shipper Name => Added on 23-July-2021 - To Avoid false positive cases
            //                    if (!string.IsNullOrEmpty(ShipperDetailsName))
            //                    {
            //                        objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ShipperDetailsName), Module1.func_RemoveSpecialCharacter(Name));
            //                        if (objRet26.PercentageMatch >= 60)
            //                        {
            //                            isNameMatch = true;
            //                        }
            //                    }

            //                    if (isAddressMatch && isCityMatch && isZipCodeMatch)
            //                    {
            //                        //// check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
            //                        //DataRow[] RowsShipperTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
            //                        //if (RowsShipperTableRecordss != null && RowsShipperTableRecordss.Length >= 1 && AddressType == "SHIPPER")
            //                        //{
            //                        //    foreach (var item in RowsShipperTableRecordss)
            //                        //    {
            //                        //        item["AddressType"] = "";
            //                        //    }
            //                        //}
            //                        AddressType = "SHIPPER";
            //                        isShipperTag = true;
            //                    }
            //                    //else if (isCityMatch && isZipCodeMatch)
            //                    else if (isCityMatch && isZipCodeMatch && isNameMatch) // Added on 23-July-2021 => To Avoid tagging issue
            //                    {
            //                        //if (Module1.TableRecordss.Select(" City ='" + ShipperDetailsCity + "' AND Zip like '" + ShipperDetailsZip + "%' ").Length == 1)
            //                        //{
            //                        //// check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
            //                        //DataRow[] RowsShipperTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
            //                        //if (RowsShipperTableRecordss != null && RowsShipperTableRecordss.Length >= 1 && AddressType == "SHIPPER")
            //                        //{
            //                        //    foreach (var item in RowsShipperTableRecordss)
            //                        //    {
            //                        //        item["AddressType"] = "";
            //                        //    }
            //                        //}
            //                        AddressType = "SHIPPER";
            //                        isShipperTag = true;
            //                        // }
            //                    }
            //                    else if (isAddressMatch && isNameMatch)
            //                    {
            //                        AddressType = "SHIPPER";
            //                        isShipperTag = true;
            //                    }
            //                }
            //                if (isConsigneeTag == false)
            //                {
            //                    if (!string.IsNullOrEmpty(ConsigneeDetailsAddressLine))
            //                    {
            //                        //objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ConsigneeDetailsAddressLine), Module1.func_RemoveSpecialCharacter(AddressLine1));
            //                        //if (objRet26.PercentageMatch >= 50)
            //                        //{
            //                        //    isAddressMatchConsignee = true;
            //                        //}
            //                        AddressLine1 = clsSearchCustomer.ReplaceAddressAbbreviation(AddressLine1.Trim(), Module1.dtAbbreviationTable);
            //                        ConsigneeDetailsAddressLine = clsSearchCustomer.ReplaceAddressAbbreviation(ConsigneeDetailsAddressLine.Trim(), Module1.dtAbbreviationTable);

            //                        var arrConsigneeDetailsAddressLine = ConsigneeDetailsAddressLine.Split(' ');
            //                        var arrAddressLine1 = AddressLine1.Split(' ');
            //                        int ConsigneeMatchedCount = arrConsigneeDetailsAddressLine.Count(x => arrAddressLine1.Contains(x));
            //                        if (ConsigneeMatchedCount >= 1 && arrConsigneeDetailsAddressLine.Length >= 1)
            //                        {
            //                            isAddressMatchConsignee = true;
            //                        }
            //                    }
            //                    // City 
            //                    if (!string.IsNullOrEmpty(ConsigneeDetailsCity))
            //                    {
            //                        //objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ConsigneeDetailsCity), Module1.func_RemoveSpecialCharacter(City));
            //                        objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(City), Module1.func_RemoveSpecialCharacter(ConsigneeDetailsCity));
            //                        if (objRet26.PercentageMatch >= 80)
            //                        {
            //                            isCityMatchConsignee = true;
            //                        }
            //                    }
            //                    // Zip 
            //                    if (!string.IsNullOrEmpty(ConsigneeDetailsZip))
            //                    {
            //                        objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ConsigneeDetailsZip), Module1.func_RemoveSpecialCharacter(ZipCode));
            //                        if (objRet26.PercentageMatch >= 80)
            //                        {
            //                            isZipCodeMatchConsignee = true;
            //                        }
            //                    }
            //                    // Consignee Name => Added on 23-July-2021 - To Avoid false positive cases
            //                    if (!string.IsNullOrEmpty(ConsigneeDetailsName))
            //                    {
            //                        objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ConsigneeDetailsName), Module1.func_RemoveSpecialCharacter(Name));
            //                        if (objRet26.PercentageMatch >= 60)
            //                        {
            //                            isNameMatchConsignee = true;
            //                        }
            //                    }
            //                    if (isAddressMatchConsignee && isCityMatchConsignee && isZipCodeMatchConsignee)
            //                    {
            //                        // check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
            //                        //DataRow[] RowsConsigneeTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
            //                        //if (RowsConsigneeTableRecordss != null && RowsConsigneeTableRecordss.Length >= 1 && AddressType == "CONSIGNEE")
            //                        //{
            //                        //    foreach (var item in RowsConsigneeTableRecordss)
            //                        //    {
            //                        //        item["AddressType"] = "";
            //                        //    }
            //                        //}
            //                        AddressType = "CONSIGNEE";
            //                        isConsigneeTag = true;
            //                        ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T"; // Added on 6-July-2021 => Tag consignee only when if we found keyword or matched with Pick up DB
            //                    }
            //                    //else if (isCityMatchConsignee && isZipCodeMatchConsignee)
            //                    else if (isCityMatchConsignee && isZipCodeMatchConsignee && isNameMatchConsignee) // Added on 23-July-2021 - To Avoid false positive cases
            //                    {
            //                        //if (Module1.TableRecordss.Select(" City ='" + ConsigneeDetailsCity + "' AND Zip like '" + ConsigneeDetailsZip + "%' ").Length == 1)
            //                        //{
            //                        //// check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
            //                        //DataRow[] RowsConsigneeTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
            //                        //if (RowsConsigneeTableRecordss != null && RowsConsigneeTableRecordss.Length >= 1 && AddressType == "CONSIGNEE")
            //                        //{
            //                        //    foreach (var item in RowsConsigneeTableRecordss)
            //                        //    {
            //                        //        item["AddressType"] = "";
            //                        //    }
            //                        //}
            //                        AddressType = "CONSIGNEE";
            //                        isConsigneeTag = true;
            //                        ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T"; // Added on 6-July-2021 => Tag consignee only when if we found keyword or matched with Pick up DB
            //                        //}
            //                    }
            //                    else if (isAddressMatchConsignee && isNameMatchConsignee)
            //                    {
            //                        AddressType = "CONSIGNEE";
            //                        isConsigneeTag = true;
            //                        ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T";
            //                    }

            //                }
            //                if (isBillToTag == false && isConsigneeTag == false && isShipperTag == false)
            //                {
            //                    if (!string.IsNullOrEmpty(BillToDetailsAddressLine))
            //                    {
            //                        //    //objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(BillToDetailsAddressLine), Module1.func_RemoveSpecialCharacter(AddressLine1));
            //                        //    //if (objRet26.PercentageMatch >= 50)
            //                        //    //{
            //                        //    //    isAddressMatchBillTo = true;
            //                        //    //}
            //                        AddressLine1 = clsSearchCustomer.ReplaceAddressAbbreviation(AddressLine1.Trim(), Module1.dtAbbreviationTable);
            //                        BillToDetailsAddressLine = clsSearchCustomer.ReplaceAddressAbbreviation(BillToDetailsAddressLine.Trim(), Module1.dtAbbreviationTable);

            //                        var arrBillToDetailsAddressLine = BillToDetailsAddressLine.Split(' ');
            //                        var arrAddressLine1 = AddressLine1.Split(' ');
            //                        int BillToMatchedCount = arrBillToDetailsAddressLine.Count(x => arrAddressLine1.Contains(x));
            //                        if (BillToMatchedCount >= 1 && arrBillToDetailsAddressLine.Length >= 1)
            //                        {
            //                            isAddressMatchBillTo = true;
            //                        }
            //                    }

            //                    // City 
            //                    if (!string.IsNullOrEmpty(BillToDetailsCity))
            //                    {
            //                        objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(BillToDetailsCity), Module1.func_RemoveSpecialCharacter(City));
            //                        if (objRet26.PercentageMatch >= 80)
            //                        {
            //                            isCityMatchBillTo = true;
            //                        }
            //                    }
            //                    // Zip 
            //                    if (!string.IsNullOrEmpty(BillToDetailsZip))
            //                    {
            //                        objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(BillToDetailsZip), Module1.func_RemoveSpecialCharacter(ZipCode));
            //                        if (objRet26.PercentageMatch >= 80)
            //                        {
            //                            isZipCodeMatchBillTo = true;
            //                        }
            //                    }
            //                    // BillTo Name => Added on 23-July-2021 - To Avoid false positive cases 
            //                    if (!string.IsNullOrEmpty(BillToDetailsName))
            //                    {
            //                        objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(BillToDetailsName), Module1.func_RemoveSpecialCharacter(Name));
            //                        if (objRet26.PercentageMatch >= 60)
            //                        {
            //                            isNameMatchBillTo = true;
            //                        }
            //                    }
            //                    if (isAddressMatchBillTo && isCityMatchBillTo && isZipCodeMatchBillTo)
            //                    {
            //                        //// check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
            //                        //DataRow[] RowsThirdPartyTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
            //                        //if (RowsThirdPartyTableRecordss != null && RowsThirdPartyTableRecordss.Length >= 1 && AddressType == "3RD PARTY")
            //                        //{
            //                        //    foreach (var item in RowsThirdPartyTableRecordss)
            //                        //    {
            //                        //        item["AddressType"] = "";
            //                        //    }
            //                        //}
            //                        AddressType = "3RD PARTY";
            //                        isBillToTag = true;
            //                    }
            //                    //else if (isCityMatchBillTo && isZipCodeMatchBillTo)
            //                    else if (isCityMatchBillTo && isZipCodeMatchBillTo && isNameMatchBillTo)
            //                    {
            //                        //if (Module1.TableRecordss.Select(" City ='" + BillToDetailsCity + "' AND Zip like '" + ConsigneeDetailsZip + "%' ").Length == 1)
            //                        //{
            //                        //// check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
            //                        //DataRow[] RowsThirdPartyTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
            //                        //if (RowsThirdPartyTableRecordss != null && RowsThirdPartyTableRecordss.Length >= 1 && AddressType == "3RD PARTY")
            //                        //{
            //                        //    foreach (var item in RowsThirdPartyTableRecordss)
            //                        //    {
            //                        //        item["AddressType"] = "";
            //                        //    }
            //                        //}
            //                        AddressType = "3RD PARTY";
            //                        isBillToTag = true;
            //                        //}
            //                    }
            //                    else if (isAddressMatchBillTo && isNameMatchBillTo)
            //                    {
            //                        AddressType = "3RD PARTY";
            //                        isBillToTag = true;
            //                    }
            //                }
            //                #endregion
            //                // BillTo Tag => if we found 3 reords out of 1 is remain without tagging then by default consider it as Bill To
            //                //if (Module1.TableRecordss.Rows.Count >= 3 && AddressType == "" && isBillToTag == false && isShipperTag == true && isConsigneeTag == true)
            //                //{
            //                //    AddressType = "3RD PARTY";
            //                //    isBillToTag = true;
            //                //}
            //            }

            //            DataRow[] RowsShipperTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
            //            if (RowsShipperTableRecordss != null && RowsShipperTableRecordss.Length >= 1 && AddressType == "SHIPPER")
            //            {
            //                AddressType = "";
            //                //foreach (var item in RowsShipperTableRecordss)
            //                //{
            //                //    item["AddressType"] = "";
            //                //}
            //            }

            //            DataRow[] RowsConsigneeTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
            //            if (RowsConsigneeTableRecordss != null && RowsConsigneeTableRecordss.Length >= 1 && AddressType == "CONSIGNEE")
            //            {
            //                AddressType = "";
            //                //foreach (var item in RowsConsigneeTableRecordss)
            //                //{
            //                //    item["AddressType"] = "";
            //                //}
            //            }

            //            DataRow[] RowsThirdPartyTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
            //            if (RowsThirdPartyTableRecordss != null && RowsThirdPartyTableRecordss.Length >= 1 && AddressType == "3RD PARTY")
            //            {
            //                AddressType = "";
            //                //foreach (var item in RowsThirdPartyTableRecordss)
            //                //{
            //                //    item["AddressType"] = "";
            //                //}
            //            }

            //            if (AddressType == "SHIPPER")
            //            {
            //                LineNo = 1;
            //            }
            //            else if (AddressType == "CONSIGNEE")
            //            {
            //                LineNo = 2;
            //            }
            //            else if (AddressType == "3RD PARTY")
            //            {
            //                LineNo = 3;
            //            }

            //            rowFinal["LineNo"] = LineNo;

            //            rowFinal["AddressType"] = AddressType;
            //            Module1.TableRecordssFinal.Rows.Add(rowFinal);
            //        }
            //        // array for loop 

            //    }
            //    #region Tag shipper/consignee/thirdparty against left one record in the TableRecordssFinal => region commented on 23-July-2021 to avoid false positive cases
            //    // Step 1 :- Identify tagging 
            //    //bool isShipperFinal = false;
            //    //bool isConsigneeFinal = false;
            //    //bool isThirdPartyFinal = false;
            //    //DataRow[] RowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
            //    //if (RowsShipper.Length > 0)
            //    //{
            //    //    isShipperFinal = true;
            //    //}
            //    //DataRow[] RowsConsignee = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
            //    //if (RowsConsignee.Length > 0)
            //    //{
            //    //    isConsigneeFinal = true;
            //    //}
            //    //DataRow[] RowsThirdParty = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
            //    //if (RowsThirdParty.Length > 0)
            //    //{
            //    //    isThirdPartyFinal = true;
            //    //}
            //    //DataRow[] RowsBlank = Module1.TableRecordssFinal.Select("[AddressType] = ''");
            //    //// Step 2 :- if we get only 3 records out of 2 are alreday tagged then tag remaining one it might be shipper/consignee/thirdparty   
            //    //if (!isShipperFinal || !isConsigneeFinal || !isThirdPartyFinal)
            //    //{
            //    //    // Step 2_1 :- check if Shipper and consignee tagged then mark 3rd one as ThirdParty
            //    //    if (RowsShipper.Length == 1 && RowsConsignee.Length == 1 && RowsBlank.Length == 1)
            //    //    {
            //    //        if (isShipperFinal && isConsigneeFinal && !isThirdPartyFinal)
            //    //        {
            //    //            if (RowsBlank != null && RowsBlank.Length == 1)
            //    //            {
            //    //                foreach (var item in RowsBlank)
            //    //                {
            //    //                    item["AddressType"] = "3RD PARTY";
            //    //                    item["LineNo"] = 3;
            //    //                }
            //    //            }
            //    //        }
            //    //    }
            //    //    // Step 2_2 :- check if Shipper and ThirdParty  tagged then mark 3rd one as consignee
            //    //    if (RowsShipper.Length == 1 && RowsThirdParty.Length == 1 && RowsBlank.Length == 1)
            //    //    {
            //    //        if (isShipperFinal && !isConsigneeFinal && isThirdPartyFinal)
            //    //        {
            //    //            if (RowsBlank != null)
            //    //            {
            //    //                foreach (var item in RowsBlank)
            //    //                {
            //    //                    // item["AddressType"] = "SHIPPER";
            //    //                    item["AddressType"] = "CONSIGNEE";
            //    //                    item["LineNo"] = 2;
            //    //                }
            //    //            }
            //    //        }
            //    //    }
            //    //    // Step 2_3 :- check if Consignee and ThirdParty  tagged then mark 3rd one as shipper
            //    //    if (RowsConsignee.Length == 1 && RowsThirdParty.Length == 1 && RowsBlank.Length == 1)
            //    //    {
            //    //        if (!isShipperFinal && isConsigneeFinal && isThirdPartyFinal)
            //    //        {
            //    //            if (RowsBlank != null)
            //    //            {
            //    //                foreach (var item in RowsBlank)
            //    //                {
            //    //                    item["AddressType"] = "SHIPPER";
            //    //                    //item["AddressType"] = "CONSIGNEE";
            //    //                    item["LineNo"] = 1;
            //    //                }
            //    //            }
            //    //        }
            //    //    }
            //    //}

            //    //Module1.TableRecordssFinal.DefaultView.Sort = "LineNo";
            //    //Module1.TableRecordssFinal = Module1.TableRecordssFinal.DefaultView.ToTable();




            //    //if (Module1.TableRecordssFinal.Rows.Count > 0)
            //    //{
            //    //    DataRow RowsBlank1 = Module1.TableRecordssFinal.Rows[0];

            //    //    if (RowsBlank1["AddressType"] == null || Convert.ToString(RowsBlank1["AddressType"]).Trim() == "")
            //    //    {
            //    //        if (Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'").Length == 0)
            //    //        {
            //    //            try
            //    //            {
            //    //                clsCnCWord _cncState = (clsCnCWord)RowsBlank1["cncState"];
            //    //                if ((Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.50)) > _cncState.Top)
            //    //                {
            //    //                    RowsBlank1["AddressType"] = "SHIPPER";
            //    //                }
            //    //            }
            //    //            catch (Exception ex)
            //    //            {
            //    //            }
            //    //        }
            //    //    }
            //    //}

            //    ////bool bConsig = false; 
            //    //if (Module1.TableRecordssFinal.Rows.Count > 1)
            //    //{
            //    //    DataRow RowsBlank1 = Module1.TableRecordssFinal.Rows[1];
            //    //    if (RowsBlank1["AddressType"] == null || Convert.ToString(RowsBlank1["AddressType"]).Trim() == "")
            //    //    {
            //    //        if (Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'").Length == 0)
            //    //        {
            //    //            try
            //    //            {
            //    //                clsCnCWord _cncState = (clsCnCWord)RowsBlank1["cncState"];
            //    //                if ((Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.75)) > _cncState.Top)
            //    //                {
            //    //                    RowsBlank1["AddressType"] = "CONSIGNEE";
            //    //                }
            //    //                else
            //    //                {
            //    //                    RowsBlank1["AddressType"] = "3RD PARTY";
            //    //                    //bConsig = true; ;
            //    //                }
            //    //            }
            //    //            catch (Exception ex)
            //    //            {
            //    //            }

            //    //        }
            //    //    }
            //    //}

            //    //if (Module1.TableRecordssFinal.Rows.Count > 2)
            //    //{
            //    //    DataRow RowsBlank1 = Module1.TableRecordssFinal.Rows[2];
            //    //    if (RowsBlank1["AddressType"] == null || Convert.ToString(RowsBlank1["AddressType"]).Trim() == "")
            //    //    {
            //    //        if (Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'").Length == 0)
            //    //        {
            //    //            RowsBlank1["AddressType"] = "3RD PARTY";
            //    //        }
            //    //    }
            //    //}


            //    ////********* if 1st row have an consignee and 2nd row blank than tagg it as bill to

            //    //if (Module1.TableRecordssFinal.Rows.Count == 2)
            //    //{
            //    //    DataRow RowsBlank1 = Module1.TableRecordssFinal.Rows[0];
            //    //    if (Convert.ToString(RowsBlank1["AddressType"]).Trim() == "CONSIGNEE")
            //    //    {
            //    //        DataRow RowsBlank1_BillTo = Module1.TableRecordssFinal.Rows[1];
            //    //        if (RowsBlank1_BillTo["AddressType"] == null || Convert.ToString(RowsBlank1_BillTo["AddressType"]).Trim() == "")
            //    //        {
            //    //            //try
            //    //            //{
            //    //            //    clsCnCWord _cncState = (clsCnCWord)RowsBlank1["cncState"];
            //    //            //    if (_cncState.Top > (Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.50)))
            //    //            //    {
            //    //            RowsBlank1_BillTo["AddressType"] = "3RD PARTY";
            //    //            //    }
            //    //            //}
            //    //            //catch (Exception ex)
            //    //            //{
            //    //            //}

            //    //        }
            //    //    }
            //    //}
            //    ////*******************************************

            //    #endregion
            //}
            #endregion

            #region Added on 30-July-2021
            if (Module1.TableRecordss.Rows.Count > 0)
            {
                Regex rxOnlyAlphaNumeric = new Regex("^(?=.*[0-9])(?=.*[a-zA-Z])([a-zA-Z0-9]+)$"); // Only alphanumeric
                Regex rxNumeric = new Regex("^[0-9]+$");
                Regex rxAlphabet = new Regex("^[a-zA-Z]+$");
                string State = String.Empty;
                string ZipCode = String.Empty;
                string City = string.Empty;
                string AddressLine1 = string.Empty;
                string AddressLine2 = string.Empty;
                string Name = string.Empty;
                string AddressType = "";
                string MergeWords = string.Empty;
                int LineNo = 0;

                //clsCnCWord clcncMergeWords = new clsCnCWord();
                //clsCnCWord cncName = new clsCnCWord();
                //clsCnCWord cncAddress = new clsCnCWord();
                //clsCnCWord cncCityStateZip = new clsCnCWord();
                //clsCnCWord cncTelPhone = new clsCnCWord();
                //clsCnCWord cncAddressType = new clsCnCWord();
                // Check shipper details
                string ShipperDetailsAddressLine = string.Empty;
                string ShipperDetailsCity = string.Empty;
                string ShipperDetailsState = string.Empty;
                string ShipperDetailsZip = string.Empty;
                string ConsigneeDetailsAddressLine = string.Empty;
                string ConsigneeDetailsCity = string.Empty;
                string ConsigneeDetailsState = string.Empty;
                string ConsigneeDetailsZip = string.Empty;
                string BillToDetailsAddressLine = string.Empty;
                string BillToDetailsCity = string.Empty;
                string BillToDetailsState = string.Empty;
                string BillToDetailsZip = string.Empty;
                // Added on 23-July-2021 => To fix wrong tagging issue
                string ShipperDetailsName = string.Empty;
                string ConsigneeDetailsName = string.Empty;
                string BillToDetailsName = string.Empty;
                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSET2_PickupShipperDetails"))
                {
                    DataTable dtKeys = new DataTable();
                    dtKeys = (DataTable)iQPUBLIC.PublicComponents.htMyVariable["BOLSET2_PickupShipperDetails"];
                    if (dtKeys != null && dtKeys.Rows.Count > 0)
                    {
                        ShipperDetailsAddressLine = dtKeys.Rows[0]["PASA1"].ToString().Trim() + dtKeys.Rows[0]["PASA2"].ToString().Trim();
                        ShipperDetailsCity = dtKeys.Rows[0]["PASCT"].ToString().Trim();
                        ShipperDetailsCity = ShipperDetailsCity.Replace("'", ""); // Added on 10-March-2020 => To handle exception in the query when city with an apostrophe in it	
                        ShipperDetailsZip = dtKeys.Rows[0]["PASZIP"].ToString().Trim();
                        ShipperDetailsName = dtKeys.Rows[0]["PASNM"].ToString().Trim(); // Added on 23-July-2021 => To fix wrong tagging issue
                        // Consignee
                        ConsigneeDetailsAddressLine = dtKeys.Rows[0]["ConsigneeAddress1"].ToString().Trim() + dtKeys.Rows[0]["ConsigneeAddress2"].ToString().Trim();
                        ConsigneeDetailsCity = dtKeys.Rows[0]["ConsigneeCity"].ToString().Trim();
                        ConsigneeDetailsCity = ConsigneeDetailsCity.Replace("'", ""); // Added on 10-March-2020 => To handle exception in the query when city with an apostrophe in it
                        ConsigneeDetailsZip = dtKeys.Rows[0]["ConsigneeZip"].ToString().Trim();
                        ConsigneeDetailsName = dtKeys.Rows[0]["ConsigneeName"].ToString().Trim(); // Added on 23-July-2021 => To fix wrong tagging issue
                        // BillTo
                        BillToDetailsAddressLine = dtKeys.Rows[0]["BillToAddress1"].ToString().Trim() + dtKeys.Rows[0]["BillToAddress2"].ToString().Trim();
                        BillToDetailsCity = dtKeys.Rows[0]["BillToCity"].ToString().Trim();
                        BillToDetailsCity = BillToDetailsCity.Replace("'", ""); // Added on 10-March-2020 => To handle exception in the query when city with an apostrophe in it
                        BillToDetailsZip = dtKeys.Rows[0]["BillToZipCode"].ToString().Trim();
                        BillToDetailsName = dtKeys.Rows[0]["BillToName"].ToString().Trim(); // Added on 23-July-2021 => To fix wrong tagging issue
                    }
                }
                //bool isShipperDouble = false;
                //bool isConsigneeDouble = false;
                //bool isThirdPartyDouble = false;
                //DataRow[] RowsShipper = Module1.TableRecordss.Select("[AddressType] = 'SHIPPER'");
                //if (RowsShipper.Length > 1)
                //{
                //    isShipperDouble = true;
                //}
                //DataRow[] RowsConsignee = Module1.TableRecordss.Select("[AddressType] = 'CONSIGNEE'");
                //if (RowsConsignee.Length > 1)
                //{
                //    isConsigneeDouble = true;
                //}
                //DataRow[] RowsThirdParty = Module1.TableRecordss.Select("[AddressType] = '3RD PARTY'");
                //if (RowsThirdParty.Length > 1)
                //{
                //    isThirdPartyDouble = true;
                //}
                bool isShipperTag = false;
                bool isConsigneeTag = false;
                bool isBillToTag = false;

                //bool currShipperRow = false;
                //bool currConsigneeRow = false;
                //bool currThirdPartyRow = false;
                for (int i = 0; i < Module1.TableRecordss.Rows.Count; i++)
                {
                    bool isValidZip = false;
                    bool isValidCity = false;
                    bool isNOTValidState = false;
                    bool isValidAddressLine1 = false;
                    clsCnCWord clcncMergeWords = new clsCnCWord();
                    clsCnCWord cncName = new clsCnCWord();
                    clsCnCWord cncAddress = new clsCnCWord();
                    clsCnCWord cncCity = new clsCnCWord();
                    clsCnCWord cncState = new clsCnCWord();
                    clsCnCWord cncZip = new clsCnCWord();
                    clsCnCWord cncCityStateZip = new clsCnCWord();
                    clsCnCWord cncTelPhone = new clsCnCWord();
                    clsCnCWord cncAddressType = new clsCnCWord();
                    DataRow rowFinal = Module1.TableRecordssFinal.NewRow();

                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["Sate"]))
                    {
                        State = Module1.TableRecordss.Rows[i]["Sate"].ToString();
                    }
                    else
                    {
                        State = "";
                    }
                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["Zip"]))
                    {
                        ZipCode = Module1.TableRecordss.Rows[i]["Zip"].ToString();
                    }
                    else
                    {
                        ZipCode = "";
                    }

                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["City"]))
                    {
                        City = Module1.TableRecordss.Rows[i]["City"].ToString();
                    }
                    else
                    {
                        City = "";
                    }
                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["AddressLine1"]))
                    {
                        AddressLine1 = Module1.TableRecordss.Rows[i]["AddressLine1"].ToString();
                    }
                    else
                    {
                        AddressLine1 = "";
                    }
                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["AddressLine2"]))
                    {
                        AddressLine2 = Module1.TableRecordss.Rows[i]["AddressLine2"].ToString();
                    }
                    else
                    {
                        AddressLine2 = "";
                    }
                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["Name"]))
                    {
                        Name = Module1.TableRecordss.Rows[i]["Name"].ToString();
                    }
                    else
                    {
                        Name = "";
                    }


                    //***************commanted on 28/07/21 RajeshB
                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["AddressType"]))
                    {
                        //AddressType = Module1.TableRecordss.Rows[i]["AddressType"].ToString();
                        cncAddressType = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncAddressType"];
                    }
                    //else
                    //{
                    //    AddressType = "";
                    //    //cncAddressType = ;
                    //}
                    AddressType = ""; //Added on 28/07/21 RajeshB
                    //*****************************************





                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["MergeWords"]))
                    {
                        MergeWords = Module1.TableRecordss.Rows[i]["MergeWords"].ToString();
                        clcncMergeWords = (clsCnCWord)Module1.TableRecordss.Rows[i]["clcncMergeWords"];

                        cncCityStateZip = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncCityStateZip"];
                    }
                    else
                    {
                        MergeWords = "";
                    }

                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncName"]))
                    {
                        cncName = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncName"];
                    }

                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncAddress"]))
                    {
                        cncAddress = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncAddress"];
                    }
                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncCity"]))
                    {
                        cncCity = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncCity"];
                    }
                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncState"]))
                    {
                        cncState = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncState"];
                    }
                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncZip"]))
                    {
                        cncZip = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncZip"];
                    }

                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["cncTelPhone"]))
                    {
                        cncTelPhone = (clsCnCWord)Module1.TableRecordss.Rows[i]["cncTelPhone"];
                    }
                    if (!DBNull.Value.Equals(Module1.TableRecordss.Rows[i]["LineNo"]))
                    {
                        LineNo = Convert.ToInt32(Module1.TableRecordss.Rows[i]["LineNo"]);
                    }
                    else
                    {
                        LineNo = 0;
                    }


                    if (!string.IsNullOrEmpty(State))
                    {
                        if (State == "ID:")
                        {
                            isNOTValidState = true;
                        }
                    }


                    if (!string.IsNullOrEmpty(ZipCode))
                    {
                        // check the valid zip code eg:- 43125-1301  then pick the first part of the zip
                        // below code commented on 21-Jan-2021 -- start
                        //if (ZipCode.Contains('-') == true)
                        //{
                        //    string[] ZipValue = ZipCode.Split('-');
                        //    if (ZipValue[0].Length <= 5)
                        //    {
                        //        ZipCode = ZipValue[0];
                        //    }
                        //}
                        // below code commented on 21-Jan-2021 -- end
                        // Code Added on 21-Jan-2021 -- start
                        if (Module1.func_RemoveSpecialCharacter(ZipCode).Length >= 5)
                        {
                            ZipCode = ZipCode.Substring(0, 5);
                        }
                        // Code Added on 21-Jan-2021 -- end
                        if ((Regex.IsMatch(Module1.func_RemoveSpecialCharacter(ZipCode).Trim(), rxNumeric.ToString()) == true || Regex.IsMatch(Module1.func_RemoveSpecialCharacter(ZipCode).Trim(), rxOnlyAlphaNumeric.ToString()) == true) && Module1.func_RemoveSpecialCharacter(ZipCode).Length >= 3)
                        {

                            isValidZip = true;

                            if (ZipCode.Contains('.') == true)
                            {
                                isValidZip = false;
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(City))
                    {
                        if ((Regex.IsMatch(Module1.func_RemoveSpecialCharacter(City).Trim(), rxAlphabet.ToString()) == true || Regex.IsMatch(Module1.func_RemoveSpecialCharacter(City).Trim(), rxOnlyAlphaNumeric.ToString()) == true) && Module1.func_RemoveSpecialCharacter(City).Length >= 3)
                        {
                            isValidCity = true;
                        }
                    }

                    if (!string.IsNullOrEmpty(AddressLine1))
                    {
                        // check the abbrivation in the AddressLine1 and check numeric with two capital Alphabet
                        if (Abbrivation.Any(AddressLine1.ToUpper().Contains))
                        {
                            isValidAddressLine1 = true;
                        }
                    }

                    if ((isValidZip && isValidCity && !isNOTValidState))
                    {
                        rowFinal["Sate"] = State;
                        rowFinal["Zip"] = ZipCode;
                        rowFinal["City"] = City;

                        if (AddressLine1 != null)
                        {
                            AddressLine1 = Module1.func_RemoveKeywordsNameAddress(AddressLine1);
                        }
                        rowFinal["AddressLine1"] = AddressLine1;
                        rowFinal["AddressLine2"] = AddressLine2;

                        if (Name != null)
                        {
                            //Name = Module1.func_RemoveKeywordsNameAddress(Name);
                            Name = Module1.func_RemoveKeywordsforAllName(Name);
                            //**********08FEB21
                            Name = Regex.Replace(Name, "^[^a-zA-Z0-9%]+", "");
                            Name = Regex.Replace(Name, "[^a-zA-Z0-9)]+$", "");
                        }
                        rowFinal["Name"] = Name;

                        rowFinal["MergeWords"] = MergeWords;
                        rowFinal["clcncMergeWords"] = clcncMergeWords;
                        if (cncName != null)
                        {
                            if (cncName.strWord != null)
                            {
                                // cncName.strWord = Module1.func_RemoveKeywordsNameAddress(cncName.strWord);
                                cncName.strWord = Module1.func_RemoveKeywordsforAllName(cncName.strWord);

                                //**********08FEB21
                                cncName.strWord = Regex.Replace(cncName.strWord, "^[^a-zA-Z0-9%]+", "");
                                cncName.strWord = Regex.Replace(cncName.strWord, "[^a-zA-Z0-9)]+$", "");
                            }
                        }
                        rowFinal["cncName"] = cncName;
                        rowFinal["cncCity"] = cncCity;
                        rowFinal["cncState"] = cncState;
                        if (cncZip != null)
                        {
                            if (cncZip.strWord != null)
                            {
                                // check the valid zip code eg:- 43125-1301  then pick the first part of the zip
                                // below code is commnted on 21-Jan-2021 -- start
                                //if (cncZip.strWord.Contains('-') == true)
                                //{
                                //    string[] ZipValue = ZipCode.Split('-');
                                //    if (ZipValue[0].Length <= 5)
                                //    {
                                //        cncZip.strWord = ZipValue[0];
                                //    }
                                //}
                                // below code is commnted on 21-Jan-2021 -- end
                                // Code Added on 21-Jan-2021 -- start
                                if (Module1.func_RemoveSpecialCharacter(cncZip.strWord).Length >= 5)
                                {
                                    cncZip.strWord = cncZip.strWord.Substring(0, 5);
                                }
                                // Code Added on 21-Jan-2021 -- end
                            }
                        }
                        rowFinal["cncZip"] = cncZip;
                        if (cncAddress != null)
                        {
                            if (cncAddress.strWord != null)
                            {
                                cncAddress.strWord = Module1.func_RemoveKeywordsNameAddress(cncAddress.strWord);

                                //**********08FEB21
                                cncAddress.strWord = Regex.Replace(cncAddress.strWord, "^[^a-zA-Z0-9%]+", "");
                                cncAddress.strWord = Regex.Replace(cncAddress.strWord, "[^a-zA-Z0-9)]+$", "");
                            }
                        }
                        rowFinal["cncAddress"] = cncAddress;
                        rowFinal["cncCityStateZip"] = cncCityStateZip;
                        if (cncTelPhone != null)
                        {
                            if (cncTelPhone.strWord != null)
                            {
                                string str_DataTel = cncTelPhone.strWord;
                                string Regxtel = "[^0-9]";
                                cncTelPhone.strWord = Regex.Replace(str_DataTel, Regxtel, "", RegexOptions.IgnoreCase);
                            }
                        }
                        rowFinal["cncTelPhone"] = cncTelPhone;
                        rowFinal["cncAddressType"] = cncAddressType;

                        // rowFinal["LineNo"] = LineNo;

                        //if (AddressType == "SHIPPER")
                        //{
                        //    currShipperRow = true;
                        //}
                        //else if (AddressType == "CONSIGNEE")
                        //{
                        //    currConsigneeRow = true;
                        //}
                        //else if (AddressType == "3RD PARTY")
                        //{
                        //    currThirdPartyRow = true;
                        //}

                        if (clsF27.clsModule.clsPickupValuesFRP001RAT != "Y")
                        {
                            // If AddressType is not tag as shipper,consignee or bill to then apply ping pong with the captured data
                            #region Ping-pong using pick up => FRP001
                            //if (AddressType == "" || isShipperDouble == true || isConsigneeDouble == true || isThirdPartyDouble == true)
                            //{
                            bool isAddressMatch = false;
                            bool isCityMatch = false;
                            bool isZipCodeMatch = false;

                            bool isAddressMatchConsignee = false;
                            bool isCityMatchConsignee = false;
                            bool isZipCodeMatchConsignee = false;

                            bool isAddressMatchBillTo = false;
                            bool isCityMatchBillTo = false;
                            bool isZipCodeMatchBillTo = false;

                            // Added on 23-July-2021 => To avoid wrong tagging
                            bool isNameMatch = false;
                            bool isNameMatchConsignee = false;
                            bool isNameMatchBillTo = false;
                            if (isShipperTag == false)
                            {
                                if (!string.IsNullOrEmpty(ShipperDetailsAddressLine))
                                {
                                    //objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ShipperDetailsAddressLine), Module1.func_RemoveSpecialCharacter(AddressLine1));
                                    //if (objRet26.PercentageMatch >= 50)
                                    //{
                                    //    isAddressMatch = true;
                                    //}

                                    AddressLine1 = clsSearchCustomer.ReplaceAddressAbbreviation(AddressLine1.Trim(), Module1.dtAbbreviationTable);
                                    ShipperDetailsAddressLine = clsSearchCustomer.ReplaceAddressAbbreviation(ShipperDetailsAddressLine.Trim(), Module1.dtAbbreviationTable);

                                    var arrShipperDetailsAddressLine = ShipperDetailsAddressLine.Split(' ');
                                    var arrAddressLine1 = AddressLine1.Split(' ');
                                    int shipperMatchedCount = arrShipperDetailsAddressLine.Count(x => arrAddressLine1.Contains(x));
                                    if (shipperMatchedCount >= 1 && arrShipperDetailsAddressLine.Length >= 1)
                                    {
                                        isAddressMatch = true;
                                    }
                                }

                                // City 
                                if (!string.IsNullOrEmpty(ShipperDetailsCity))
                                {
                                    //objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ShipperDetailsCity), Module1.func_RemoveSpecialCharacter(City));
                                    objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(City), Module1.func_RemoveSpecialCharacter(ShipperDetailsCity));
                                    if (objRet26.PercentageMatch >= 80)
                                    {
                                        isCityMatch = true;
                                    }
                                }
                                // Zip 
                                if (!string.IsNullOrEmpty(ShipperDetailsZip))
                                {
                                    objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ShipperDetailsZip), Module1.func_RemoveSpecialCharacter(ZipCode));
                                    if (objRet26.PercentageMatch >= 80)
                                    {
                                        isZipCodeMatch = true;
                                    }
                                }
                                // Shipper Name => Added on 23-July-2021 - To Avoid false positive cases
                                if (!string.IsNullOrEmpty(ShipperDetailsName))
                                {
                                    objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ShipperDetailsName), Module1.func_RemoveSpecialCharacter(Name));
                                    if (objRet26.PercentageMatch >= 60)
                                    {
                                        isNameMatch = true;
                                    }
                                }

                                if (isAddressMatch && isCityMatch && isZipCodeMatch)
                                {
                                    //// check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
                                    //DataRow[] RowsShipperTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
                                    //if (RowsShipperTableRecordss != null && RowsShipperTableRecordss.Length >= 1 && AddressType == "SHIPPER")
                                    //{
                                    //    foreach (var item in RowsShipperTableRecordss)
                                    //    {
                                    //        item["AddressType"] = "";
                                    //    }
                                    //}
                                    AddressType = "SHIPPER";
                                    isShipperTag = true;
                                }
                                //else if (isCityMatch && isZipCodeMatch)
                                else if (isCityMatch && isZipCodeMatch && isNameMatch) // Added on 23-July-2021 => To Avoid tagging issue
                                {
                                    //if (Module1.TableRecordss.Select(" City ='" + ShipperDetailsCity + "' AND Zip like '" + ShipperDetailsZip + "%' ").Length == 1)
                                    //{
                                    //// check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
                                    //DataRow[] RowsShipperTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
                                    //if (RowsShipperTableRecordss != null && RowsShipperTableRecordss.Length >= 1 && AddressType == "SHIPPER")
                                    //{
                                    //    foreach (var item in RowsShipperTableRecordss)
                                    //    {
                                    //        item["AddressType"] = "";
                                    //    }
                                    //}
                                    AddressType = "SHIPPER";
                                    isShipperTag = true;
                                    // }
                                }
                                else if (isAddressMatch && isNameMatch)
                                {
                                    AddressType = "SHIPPER";
                                    isShipperTag = true;
                                }


                            }
                            if (isConsigneeTag == false)
                            {
                                if (!string.IsNullOrEmpty(ConsigneeDetailsAddressLine))
                                {
                                    //objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ConsigneeDetailsAddressLine), Module1.func_RemoveSpecialCharacter(AddressLine1));
                                    //if (objRet26.PercentageMatch >= 50)
                                    //{
                                    //    isAddressMatchConsignee = true;
                                    //}
                                    AddressLine1 = clsSearchCustomer.ReplaceAddressAbbreviation(AddressLine1.Trim(), Module1.dtAbbreviationTable);
                                    ConsigneeDetailsAddressLine = clsSearchCustomer.ReplaceAddressAbbreviation(ConsigneeDetailsAddressLine.Trim(), Module1.dtAbbreviationTable);

                                    var arrConsigneeDetailsAddressLine = ConsigneeDetailsAddressLine.Split(' ');
                                    var arrAddressLine1 = AddressLine1.Split(' ');
                                    int ConsigneeMatchedCount = arrConsigneeDetailsAddressLine.Count(x => arrAddressLine1.Contains(x));
                                    if (ConsigneeMatchedCount >= 1 && arrConsigneeDetailsAddressLine.Length >= 1)
                                    {
                                        isAddressMatchConsignee = true;
                                    }
                                }
                                // City 
                                if (!string.IsNullOrEmpty(ConsigneeDetailsCity))
                                {
                                    //objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ConsigneeDetailsCity), Module1.func_RemoveSpecialCharacter(City));
                                    objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(City), Module1.func_RemoveSpecialCharacter(ConsigneeDetailsCity));
                                    if (objRet26.PercentageMatch >= 80)
                                    {
                                        isCityMatchConsignee = true;
                                    }
                                }
                                // Zip 
                                if (!string.IsNullOrEmpty(ConsigneeDetailsZip))
                                {
                                    objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ConsigneeDetailsZip), Module1.func_RemoveSpecialCharacter(ZipCode));
                                    if (objRet26.PercentageMatch >= 80)
                                    {
                                        isZipCodeMatchConsignee = true;
                                    }
                                }
                                // Consignee Name => Added on 23-July-2021 - To Avoid false positive cases
                                if (!string.IsNullOrEmpty(ConsigneeDetailsName))
                                {
                                    objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(ConsigneeDetailsName), Module1.func_RemoveSpecialCharacter(Name));
                                    if (objRet26.PercentageMatch >= 60)
                                    {
                                        isNameMatchConsignee = true;
                                    }
                                }
                                if (isAddressMatchConsignee && isCityMatchConsignee && isZipCodeMatchConsignee)
                                {
                                    // check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
                                    //DataRow[] RowsConsigneeTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
                                    //if (RowsConsigneeTableRecordss != null && RowsConsigneeTableRecordss.Length >= 1 && AddressType == "CONSIGNEE")
                                    //{
                                    //    foreach (var item in RowsConsigneeTableRecordss)
                                    //    {
                                    //        item["AddressType"] = "";
                                    //    }
                                    //}
                                    AddressType = "CONSIGNEE";
                                    isConsigneeTag = true;
                                    ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T"; // Added on 6-July-2021 => Tag consignee only when if we found keyword or matched with Pick up DB
                                }
                                //else if (isCityMatchConsignee && isZipCodeMatchConsignee)
                                else if (isCityMatchConsignee && isZipCodeMatchConsignee && isNameMatchConsignee) // Added on 23-July-2021 - To Avoid false positive cases
                                {
                                    //if (Module1.TableRecordss.Select(" City ='" + ConsigneeDetailsCity + "' AND Zip like '" + ConsigneeDetailsZip + "%' ").Length == 1)
                                    //{
                                    //// check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
                                    //DataRow[] RowsConsigneeTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
                                    //if (RowsConsigneeTableRecordss != null && RowsConsigneeTableRecordss.Length >= 1 && AddressType == "CONSIGNEE")
                                    //{
                                    //    foreach (var item in RowsConsigneeTableRecordss)
                                    //    {
                                    //        item["AddressType"] = "";
                                    //    }
                                    //}
                                    AddressType = "CONSIGNEE";
                                    isConsigneeTag = true;
                                    ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T"; // Added on 6-July-2021 => Tag consignee only when if we found keyword or matched with Pick up DB
                                    //}
                                }
                                else if (isAddressMatchConsignee && isNameMatchConsignee)
                                {
                                    AddressType = "CONSIGNEE";
                                    isConsigneeTag = true;
                                    ClsOCRValues.ClsOCRCorrectConsigneeTagged = "T";
                                }

                            }
                            if (isBillToTag == false && isConsigneeTag == false && isShipperTag == false)
                            {
                                if (!string.IsNullOrEmpty(BillToDetailsAddressLine))
                                {
                                    //    //objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(BillToDetailsAddressLine), Module1.func_RemoveSpecialCharacter(AddressLine1));
                                    //    //if (objRet26.PercentageMatch >= 50)
                                    //    //{
                                    //    //    isAddressMatchBillTo = true;
                                    //    //}
                                    AddressLine1 = clsSearchCustomer.ReplaceAddressAbbreviation(AddressLine1.Trim(), Module1.dtAbbreviationTable);
                                    BillToDetailsAddressLine = clsSearchCustomer.ReplaceAddressAbbreviation(BillToDetailsAddressLine.Trim(), Module1.dtAbbreviationTable);

                                    var arrBillToDetailsAddressLine = BillToDetailsAddressLine.Split(' ');
                                    var arrAddressLine1 = AddressLine1.Split(' ');
                                    int BillToMatchedCount = arrBillToDetailsAddressLine.Count(x => arrAddressLine1.Contains(x));
                                    if (BillToMatchedCount >= 1 && arrBillToDetailsAddressLine.Length >= 1)
                                    {
                                        isAddressMatchBillTo = true;
                                    }
                                }

                                // City 
                                if (!string.IsNullOrEmpty(BillToDetailsCity))
                                {
                                    objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(BillToDetailsCity), Module1.func_RemoveSpecialCharacter(City));
                                    if (objRet26.PercentageMatch >= 80)
                                    {
                                        isCityMatchBillTo = true;
                                    }
                                }
                                // Zip 
                                if (!string.IsNullOrEmpty(BillToDetailsZip))
                                {
                                    objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(BillToDetailsZip), Module1.func_RemoveSpecialCharacter(ZipCode));
                                    if (objRet26.PercentageMatch >= 80)
                                    {
                                        isZipCodeMatchBillTo = true;
                                    }
                                }
                                // BillTo Name => Added on 23-July-2021 - To Avoid false positive cases 
                                if (!string.IsNullOrEmpty(BillToDetailsName))
                                {
                                    objRet26 = ObjSbr.IQSBR026(Module1.func_RemoveSpecialCharacter(BillToDetailsName), Module1.func_RemoveSpecialCharacter(Name));
                                    if (objRet26.PercentageMatch >= 60)
                                    {
                                        isNameMatchBillTo = true;
                                    }
                                }
                                if (isAddressMatchBillTo && isCityMatchBillTo && isZipCodeMatchBillTo)
                                {
                                    //// check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
                                    //DataRow[] RowsThirdPartyTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
                                    //if (RowsThirdPartyTableRecordss != null && RowsThirdPartyTableRecordss.Length >= 1 && AddressType == "3RD PARTY")
                                    //{
                                    //    foreach (var item in RowsThirdPartyTableRecordss)
                                    //    {
                                    //        item["AddressType"] = "";
                                    //    }
                                    //}
                                    AddressType = "3RD PARTY";
                                    isBillToTag = true;
                                }
                                //else if (isCityMatchBillTo && isZipCodeMatchBillTo)
                                else if (isCityMatchBillTo && isZipCodeMatchBillTo && isNameMatchBillTo)
                                {
                                    //if (Module1.TableRecordss.Select(" City ='" + BillToDetailsCity + "' AND Zip like '" + ConsigneeDetailsZip + "%' ").Length == 1)
                                    //{
                                    //// check in the table if shipper is wrongly marked previouly and we found the correct address then make address type is blank and tag for corrcet address
                                    //DataRow[] RowsThirdPartyTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
                                    //if (RowsThirdPartyTableRecordss != null && RowsThirdPartyTableRecordss.Length >= 1 && AddressType == "3RD PARTY")
                                    //{
                                    //    foreach (var item in RowsThirdPartyTableRecordss)
                                    //    {
                                    //        item["AddressType"] = "";
                                    //    }
                                    //}
                                    AddressType = "3RD PARTY";
                                    isBillToTag = true;
                                    //}
                                }
                                else if (isAddressMatchBillTo && isNameMatchBillTo)
                                {
                                    AddressType = "3RD PARTY";
                                    isBillToTag = true;
                                }
                            }
                            #endregion
                            // BillTo Tag => if we found 3 reords out of 1 is remain without tagging then by default consider it as Bill To
                            //if (Module1.TableRecordss.Rows.Count >= 3 && AddressType == "" && isBillToTag == false && isShipperTag == true && isConsigneeTag == true)
                            //{
                            //    AddressType = "3RD PARTY";
                            //    isBillToTag = true;
                            //}
                        }

                        //DataRow[] RowsShipperTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
                        //if (RowsShipperTableRecordss != null && RowsShipperTableRecordss.Length >= 1 && AddressType == "SHIPPER")
                        //{

                        //    AddressType = "";
                        //    //foreach (var item in RowsShipperTableRecordss)
                        //    //{
                        //    //    item["AddressType"] = "";
                        //    //}
                        //}

                        //DataRow[] RowsConsigneeTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
                        //if (RowsConsigneeTableRecordss != null && RowsConsigneeTableRecordss.Length >= 1 && AddressType == "CONSIGNEE")
                        //{
                        //    AddressType = "";
                        //    //foreach (var item in RowsConsigneeTableRecordss)
                        //    //{
                        //    //    item["AddressType"] = "";
                        //    //}
                        //}

                        //DataRow[] RowsThirdPartyTableRecordss = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
                        //if (RowsThirdPartyTableRecordss != null && RowsThirdPartyTableRecordss.Length >= 1 && AddressType == "3RD PARTY")
                        //{
                        //    AddressType = "";
                        //    //foreach (var item in RowsThirdPartyTableRecordss)
                        //    //{
                        //    //    item["AddressType"] = "";
                        //    //}
                        //}



                        if (AddressType == "SHIPPER")
                        {
                            LineNo = 1;
                        }
                        if (AddressType == "CONSIGNEE")
                        {
                            LineNo = 2;
                        }
                        if (AddressType == "3RD PARTY")
                        {
                            LineNo = 3;
                        }


                        if (isShipperTag == true && AddressType == "SHIPPER" || isConsigneeTag == true && AddressType == "CONSIGNEE" || isBillToTag == true && AddressType == "3RD PARTY")
                        {

                            rowFinal["LineNo"] = LineNo;

                            rowFinal["AddressType"] = AddressType;
                            Module1.TableRecordssFinal.Rows.Add(rowFinal);

                            //DataRow dr=    Module1.TableRecordss.Rows[i]["AddressType"]

                            DataRow dr = Module1.TableRecordss.Rows[i];
                            dr["AddressType"] = "";
                            Module1.TableRecordss.AcceptChanges();

                        }
                    }
                    // array for loop 

                }


                #region "If shipper/consignee/thirdparty not tagged Using Ping-pong -- use Keyword AddressType from Module1.TableRecordss "

                if (Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'").Length == 0)
                {
                    DataRow[] RowsShipper_Keyword = Module1.TableRecordss.Select("[AddressType] = 'SHIPPER'");
                    if (RowsShipper_Keyword.Length == 1)
                    {
                        //DataRow rowFinal = Module1.TableRecordssFinal.NewRow();
                        DataRow rowFinal = drKeywordAddressTag(RowsShipper_Keyword[0]);

                        if (rowFinal != null)
                        {
                            rowFinal["LineNo"] = 1;
                            rowFinal["AddressType"] = "SHIPPER";
                            Module1.TableRecordssFinal.Rows.Add(rowFinal);
                        }
                    }
                }

                if (Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'").Length == 0)
                {
                    DataRow[] RowsConsignee_Keyword = Module1.TableRecordss.Select("[AddressType] = 'CONSIGNEE'");
                    if (RowsConsignee_Keyword.Length == 1)
                    {
                        DataRow rowFinal = drKeywordAddressTag(RowsConsignee_Keyword[0]);

                        if (rowFinal != null)
                        {
                            rowFinal["LineNo"] = 2;
                            rowFinal["AddressType"] = "CONSIGNEE";
                            Module1.TableRecordssFinal.Rows.Add(rowFinal);
                        }
                    }
                }

                if (Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'").Length == 0)
                {
                    DataRow[] RowsThirdParty_Keyword = Module1.TableRecordss.Select("[AddressType] = '3RD PARTY'");
                    if (RowsThirdParty_Keyword.Length == 1)
                    {
                        try
                        {
                            // DataRow rowFinal = Module1.TableRecordssFinal.NewRow();
                            DataRow rowFinal = drKeywordAddressTag(RowsThirdParty_Keyword[0]);

                            if (rowFinal != null)
                            {
                                rowFinal["LineNo"] = 3;
                                rowFinal["AddressType"] = "3RD PARTY";
                                Module1.TableRecordssFinal.Rows.Add(rowFinal);
                            }

                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show(ex.Message);
                        }
                    }
                }

                #endregion

                #region Tag shipper/consignee/thirdparty against left one record in the TableRecordssFinal => region commented on 23-July-2021 to avoid false positive cases
                // Step 1 :- Identify tagging 
                //bool isShipperFinal = false;
                //bool isConsigneeFinal = false;
                //bool isThirdPartyFinal = false;
                //DataRow[] RowsShipper = Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'");
                //if (RowsShipper.Length > 0)
                //{
                //    isShipperFinal = true;
                //}
                //DataRow[] RowsConsignee = Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'");
                //if (RowsConsignee.Length > 0)
                //{
                //    isConsigneeFinal = true;
                //}
                //DataRow[] RowsThirdParty = Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'");
                //if (RowsThirdParty.Length > 0)
                //{
                //    isThirdPartyFinal = true;
                //}
                //DataRow[] RowsBlank = Module1.TableRecordssFinal.Select("[AddressType] = ''");
                //// Step 2 :- if we get only 3 records out of 2 are alreday tagged then tag remaining one it might be shipper/consignee/thirdparty   
                //if (!isShipperFinal || !isConsigneeFinal || !isThirdPartyFinal)
                //{
                //    // Step 2_1 :- check if Shipper and consignee tagged then mark 3rd one as ThirdParty
                //    if (RowsShipper.Length == 1 && RowsConsignee.Length == 1 && RowsBlank.Length == 1)
                //    {
                //        if (isShipperFinal && isConsigneeFinal && !isThirdPartyFinal)
                //        {
                //            if (RowsBlank != null && RowsBlank.Length == 1)
                //            {
                //                foreach (var item in RowsBlank)
                //                {
                //                    item["AddressType"] = "3RD PARTY";
                //                    item["LineNo"] = 3;
                //                }
                //            }
                //        }
                //    }
                //    // Step 2_2 :- check if Shipper and ThirdParty  tagged then mark 3rd one as consignee
                //    if (RowsShipper.Length == 1 && RowsThirdParty.Length == 1 && RowsBlank.Length == 1)
                //    {
                //        if (isShipperFinal && !isConsigneeFinal && isThirdPartyFinal)
                //        {
                //            if (RowsBlank != null)
                //            {
                //                foreach (var item in RowsBlank)
                //                {
                //                    // item["AddressType"] = "SHIPPER";
                //                    item["AddressType"] = "CONSIGNEE";
                //                    item["LineNo"] = 2;
                //                }
                //            }
                //        }
                //    }
                //    // Step 2_3 :- check if Consignee and ThirdParty  tagged then mark 3rd one as shipper
                //    if (RowsConsignee.Length == 1 && RowsThirdParty.Length == 1 && RowsBlank.Length == 1)
                //    {
                //        if (!isShipperFinal && isConsigneeFinal && isThirdPartyFinal)
                //        {
                //            if (RowsBlank != null)
                //            {
                //                foreach (var item in RowsBlank)
                //                {
                //                    item["AddressType"] = "SHIPPER";
                //                    //item["AddressType"] = "CONSIGNEE";
                //                    item["LineNo"] = 1;
                //                }
                //            }
                //        }
                //    }
                //}

                //Module1.TableRecordssFinal.DefaultView.Sort = "LineNo";
                //Module1.TableRecordssFinal = Module1.TableRecordssFinal.DefaultView.ToTable();




                //if (Module1.TableRecordssFinal.Rows.Count > 0)
                //{
                //    DataRow RowsBlank1 = Module1.TableRecordssFinal.Rows[0];

                //    if (RowsBlank1["AddressType"] == null || Convert.ToString(RowsBlank1["AddressType"]).Trim() == "")
                //    {
                //        if (Module1.TableRecordssFinal.Select("[AddressType] = 'SHIPPER'").Length == 0)
                //        {
                //            try
                //            {
                //                clsCnCWord _cncState = (clsCnCWord)RowsBlank1["cncState"];
                //                if ((Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.50)) > _cncState.Top)
                //                {
                //                    RowsBlank1["AddressType"] = "SHIPPER";
                //                }
                //            }
                //            catch (Exception ex)
                //            {
                //            }
                //        }
                //    }
                //}

                ////bool bConsig = false; 
                //if (Module1.TableRecordssFinal.Rows.Count > 1)
                //{
                //    DataRow RowsBlank1 = Module1.TableRecordssFinal.Rows[1];
                //    if (RowsBlank1["AddressType"] == null || Convert.ToString(RowsBlank1["AddressType"]).Trim() == "")
                //    {
                //        if (Module1.TableRecordssFinal.Select("[AddressType] = 'CONSIGNEE'").Length == 0)
                //        {
                //            try
                //            {
                //                clsCnCWord _cncState = (clsCnCWord)RowsBlank1["cncState"];
                //                if ((Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.75)) > _cncState.Top)
                //                {
                //                    RowsBlank1["AddressType"] = "CONSIGNEE";
                //                }
                //                else
                //                {
                //                    RowsBlank1["AddressType"] = "3RD PARTY";
                //                    //bConsig = true; ;
                //                }
                //            }
                //            catch (Exception ex)
                //            {
                //            }

                //        }
                //    }
                //}

                //if (Module1.TableRecordssFinal.Rows.Count > 2)
                //{
                //    DataRow RowsBlank1 = Module1.TableRecordssFinal.Rows[2];
                //    if (RowsBlank1["AddressType"] == null || Convert.ToString(RowsBlank1["AddressType"]).Trim() == "")
                //    {
                //        if (Module1.TableRecordssFinal.Select("[AddressType] = '3RD PARTY'").Length == 0)
                //        {
                //            RowsBlank1["AddressType"] = "3RD PARTY";
                //        }
                //    }
                //}


                ////********* if 1st row have an consignee and 2nd row blank than tagg it as bill to

                //if (Module1.TableRecordssFinal.Rows.Count == 2)
                //{
                //    DataRow RowsBlank1 = Module1.TableRecordssFinal.Rows[0];
                //    if (Convert.ToString(RowsBlank1["AddressType"]).Trim() == "CONSIGNEE")
                //    {
                //        DataRow RowsBlank1_BillTo = Module1.TableRecordssFinal.Rows[1];
                //        if (RowsBlank1_BillTo["AddressType"] == null || Convert.ToString(RowsBlank1_BillTo["AddressType"]).Trim() == "")
                //        {
                //            //try
                //            //{
                //            //    clsCnCWord _cncState = (clsCnCWord)RowsBlank1["cncState"];
                //            //    if (_cncState.Top > (Convert.ToInt32(ObjMetaData.Page[intCurrPageNumber].ImageHeight * 0.50)))
                //            //    {
                //            RowsBlank1_BillTo["AddressType"] = "3RD PARTY";
                //            //    }
                //            //}
                //            //catch (Exception ex)
                //            //{
                //            //}

                //        }
                //    }
                //}
                ////*******************************************

                #endregion
            }
            #endregion
            PossibleWords.Add(GetDefaultWord(intCurrPageNumber, "*****"));
            Module1.objRetStructF3.Words = PossibleWords;
            ConfLevel.Add(90);
            Module1.objRetStructF3.ManualConfirmation = "N";
            Module1.objRetStructF3.Status = "S";
            Module1.objRetStructF3.Flag = "Y";
            Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
            if (wordsDictionary != null)
            {
                wordsDictionary.Dispose();
                wordsDictionary = null;
            }
            if (Module1.TableRecordss != null)
            {
                Module1.TableRecordss.Dispose();
                Module1.TableRecordss = null;
            }
            return Module1.objRetStructF3;

        }
        private DataRow drKeywordAddressTag(DataRow dr)
        {
            DataRow rowFinal = Module1.TableRecordssFinal.NewRow();
            //DataRow rowFinal = null;
            try
            {
                string State = String.Empty;
                string ZipCode = String.Empty;
                string City = string.Empty;
                string AddressLine1 = string.Empty;
                string AddressLine2 = string.Empty;
                string Name = string.Empty;
                string MergeWords = string.Empty;


                Regex rxOnlyAlphaNumeric = new Regex("^(?=.*[0-9])(?=.*[a-zA-Z])([a-zA-Z0-9]+)$"); // Only alphanumeric
                Regex rxNumeric = new Regex("^[0-9]+$");
                Regex rxAlphabet = new Regex("^[a-zA-Z]+$");


                //string AddressType = "";

                int LineNo = 0;

                bool isValidZip = false;
                bool isValidCity = false;
                bool isNOTValidState = false;
                //bool isValidAddressLine1 = false;

                clsCnCWord clcncMergeWords = new clsCnCWord();
                clsCnCWord cncCityStateZip = new clsCnCWord();

                clsCnCWord cncName = new clsCnCWord();
                clsCnCWord cncAddress = new clsCnCWord();
                clsCnCWord cncCity = new clsCnCWord();
                clsCnCWord cncState = new clsCnCWord();
                clsCnCWord cncZip = new clsCnCWord();

                clsCnCWord cncTelPhone = new clsCnCWord();
                clsCnCWord cncAddressType = new clsCnCWord();
                //DataRow rowFinal = Module1.TableRecordssFinal.NewRow();

                if (!DBNull.Value.Equals(dr["Sate"]))
                {
                    State = dr["Sate"].ToString();
                }
                else
                {
                    State = "";
                }
                if (!DBNull.Value.Equals(dr["Zip"]))
                {
                    ZipCode = dr["Zip"].ToString();
                }
                else
                {
                    ZipCode = "";
                }

                if (!DBNull.Value.Equals(dr["City"]))
                {
                    City = dr["City"].ToString();
                }
                else
                {
                    City = "";
                }
                if (!DBNull.Value.Equals(dr["AddressLine1"]))
                {
                    AddressLine1 = dr["AddressLine1"].ToString();
                }
                else
                {
                    AddressLine1 = "";
                }
                if (!DBNull.Value.Equals(dr["AddressLine2"]))
                {
                    AddressLine2 = dr["AddressLine2"].ToString();
                }
                else
                {
                    AddressLine2 = "";
                }
                if (!DBNull.Value.Equals(dr["Name"]))
                {
                    Name = dr["Name"].ToString();
                }
                else
                {
                    Name = "";
                }


                //***************commanted on 28/07/21 RajeshB
                //if (!DBNull.Value.Equals(dr["AddressType"]))
                //{
                //    //AddressType = dr["AddressType"].ToString();
                //    cncAddressType = (clsCnCWord)dr["cncAddressType"];
                //}
                ////else
                ////{
                ////    AddressType = "";
                ////    //cncAddressType = ;
                ////}
                //AddressType = ""; //Added on 28/07/21 RajeshB
                //*****************************************





                if (!DBNull.Value.Equals(dr["MergeWords"]))
                {
                    MergeWords = dr["MergeWords"].ToString();
                    clcncMergeWords = (clsCnCWord)dr["clcncMergeWords"];

                    cncCityStateZip = (clsCnCWord)dr["cncCityStateZip"];
                }
                else
                {
                    MergeWords = "";
                }

                if (!DBNull.Value.Equals(dr["cncName"]))
                {
                    cncName = (clsCnCWord)dr["cncName"];
                }

                if (!DBNull.Value.Equals(dr["cncAddress"]))
                {
                    cncAddress = (clsCnCWord)dr["cncAddress"];
                }
                if (!DBNull.Value.Equals(dr["cncCity"]))
                {
                    cncCity = (clsCnCWord)dr["cncCity"];
                }
                if (!DBNull.Value.Equals(dr["cncState"]))
                {
                    cncState = (clsCnCWord)dr["cncState"];
                }
                if (!DBNull.Value.Equals(dr["cncZip"]))
                {
                    cncZip = (clsCnCWord)dr["cncZip"];
                }

                if (!DBNull.Value.Equals(dr["cncTelPhone"]))
                {
                    cncTelPhone = (clsCnCWord)dr["cncTelPhone"];
                }
                if (!DBNull.Value.Equals(dr["LineNo"]))
                {
                    LineNo = Convert.ToInt32(dr["LineNo"]);
                }
                else
                {
                    LineNo = 0;
                }


                if (!string.IsNullOrEmpty(State))
                {
                    if (State == "ID:")
                    {
                        isNOTValidState = true;
                    }
                }


                if (!string.IsNullOrEmpty(ZipCode))
                {
                    // check the valid zip code eg:- 43125-1301  then pick the first part of the zip
                    // below code commented on 21-Jan-2021 -- start
                    //if (ZipCode.Contains('-') == true)
                    //{
                    //    string[] ZipValue = ZipCode.Split('-');
                    //    if (ZipValue[0].Length <= 5)
                    //    {
                    //        ZipCode = ZipValue[0];
                    //    }
                    //}
                    // below code commented on 21-Jan-2021 -- end
                    // Code Added on 21-Jan-2021 -- start
                    if (Module1.func_RemoveSpecialCharacter(ZipCode).Length >= 5)
                    {
                        ZipCode = ZipCode.Substring(0, 5);
                    }
                    // Code Added on 21-Jan-2021 -- end
                    if ((Regex.IsMatch(Module1.func_RemoveSpecialCharacter(ZipCode).Trim(), rxNumeric.ToString()) == true || Regex.IsMatch(Module1.func_RemoveSpecialCharacter(ZipCode).Trim(), rxOnlyAlphaNumeric.ToString()) == true) && Module1.func_RemoveSpecialCharacter(ZipCode).Length >= 3)
                    {

                        isValidZip = true;

                        if (ZipCode.Contains('.') == true)
                        {
                            isValidZip = false;
                        }
                    }
                }
                if (!string.IsNullOrEmpty(City))
                {
                    if ((Regex.IsMatch(Module1.func_RemoveSpecialCharacter(City).Trim(), rxAlphabet.ToString()) == true || Regex.IsMatch(Module1.func_RemoveSpecialCharacter(City).Trim(), rxOnlyAlphaNumeric.ToString()) == true) && Module1.func_RemoveSpecialCharacter(City).Length >= 3)
                    {
                        isValidCity = true;
                    }
                }

                //if (!string.IsNullOrEmpty(AddressLine1))
                //{
                //    // check the abbrivation in the AddressLine1 and check numeric with two capital Alphabet
                //    if (Abbrivation.Any(AddressLine1.ToUpper().Contains))
                //    {
                //        isValidAddressLine1 = true;
                //    }
                //}

                if ((isValidZip && isValidCity && !isNOTValidState))
                {
                    rowFinal["Sate"] = State;
                    rowFinal["Zip"] = ZipCode;
                    rowFinal["City"] = City;

                    if (AddressLine1 != null)
                    {
                        AddressLine1 = Module1.func_RemoveKeywordsNameAddress(AddressLine1);
                    }
                    rowFinal["AddressLine1"] = AddressLine1;
                    rowFinal["AddressLine2"] = AddressLine2;

                    if (Name != null)
                    {
                        //Name = Module1.func_RemoveKeywordsNameAddress(Name);
                        Name = Module1.func_RemoveKeywordsforAllName(Name);
                        //**********08FEB21
                        Name = Regex.Replace(Name, "^[^a-zA-Z0-9%]+", "");
                        Name = Regex.Replace(Name, "[^a-zA-Z0-9)]+$", "");
                    }
                    rowFinal["Name"] = Name;

                    rowFinal["MergeWords"] = MergeWords;
                    rowFinal["clcncMergeWords"] = clcncMergeWords;
                    if (cncName != null)
                    {
                        if (cncName.strWord != null)
                        {
                            // cncName.strWord = Module1.func_RemoveKeywordsNameAddress(cncName.strWord);
                            cncName.strWord = Module1.func_RemoveKeywordsforAllName(cncName.strWord);

                            //**********08FEB21
                            cncName.strWord = Regex.Replace(cncName.strWord, "^[^a-zA-Z0-9%]+", "");
                            cncName.strWord = Regex.Replace(cncName.strWord, "[^a-zA-Z0-9)]+$", "");
                        }
                    }
                    rowFinal["cncName"] = cncName;
                    rowFinal["cncCity"] = cncCity;
                    rowFinal["cncState"] = cncState;
                    if (cncZip != null)
                    {
                        if (cncZip.strWord != null)
                        {
                            // check the valid zip code eg:- 43125-1301  then pick the first part of the zip
                            // below code is commnted on 21-Jan-2021 -- start
                            //if (cncZip.strWord.Contains('-') == true)
                            //{
                            //    string[] ZipValue = ZipCode.Split('-');
                            //    if (ZipValue[0].Length <= 5)
                            //    {
                            //        cncZip.strWord = ZipValue[0];
                            //    }
                            //}
                            // below code is commnted on 21-Jan-2021 -- end
                            // Code Added on 21-Jan-2021 -- start
                            if (Module1.func_RemoveSpecialCharacter(cncZip.strWord).Length >= 5)
                            {
                                cncZip.strWord = cncZip.strWord.Substring(0, 5);
                            }
                            // Code Added on 21-Jan-2021 -- end
                        }
                    }
                    rowFinal["cncZip"] = cncZip;
                    if (cncAddress != null)
                    {
                        if (cncAddress.strWord != null)
                        {
                            cncAddress.strWord = Module1.func_RemoveKeywordsNameAddress(cncAddress.strWord);

                            //**********08FEB21
                            cncAddress.strWord = Regex.Replace(cncAddress.strWord, "^[^a-zA-Z0-9%]+", "");
                            cncAddress.strWord = Regex.Replace(cncAddress.strWord, "[^a-zA-Z0-9)]+$", "");
                        }
                    }
                    rowFinal["cncAddress"] = cncAddress;
                    rowFinal["cncCityStateZip"] = cncCityStateZip;
                    if (cncTelPhone != null)
                    {
                        if (cncTelPhone.strWord != null)
                        {
                            string str_DataTel = cncTelPhone.strWord;
                            string Regxtel = "[^0-9]";
                            cncTelPhone.strWord = Regex.Replace(str_DataTel, Regxtel, "", RegexOptions.IgnoreCase);
                        }
                    }
                    rowFinal["cncTelPhone"] = cncTelPhone;
                    rowFinal["cncAddressType"] = cncAddressType;




                }
                else
                {
                    rowFinal = null;
                }


            }
            catch (Exception ex)
            {
                rowFinal = null;
                MessageBox.Show(ex.Message);
            }
            return rowFinal;

        }
        private clsCnCWord GetDefaultWord_StateZipCode_NewLogic(int iCurpgno, clsCnCLine clscncvalue, string strval)
        {
            clsCnCWord oRetword = new clsCnCWord();
            oRetword.X1Char = clscncvalue.LineX1Char;
            oRetword.Y1Char = clscncvalue.LineY1Char;
            oRetword.X2Char = clscncvalue.LineX2Char;
            oRetword.Y2Char = clscncvalue.LineY2Char;
            oRetword.Confidence = clscncvalue.Confidence;
            oRetword.LineNo = clscncvalue.LineNo;
            oRetword.PageNo = iCurpgno;
            oRetword.Left = clscncvalue.LineLeft;
            oRetword.Right = clscncvalue.lineRight;
            oRetword.Top = clscncvalue.LineTop;
            oRetword.Bottom = clscncvalue.lineBottom;
            oRetword.strWord = strval.Trim();
            oRetword.ConfString = "9".PadLeft(strval.Trim().Length);
            return oRetword;
        }
        private clsCnCWord GetDefaultWord_StateZipCode(int iCurpgno, clsCnCWord clscncvalue, string strval)
        {
            clsCnCWord oRetword = new clsCnCWord();
            oRetword.X1Char = clscncvalue.X1Char;
            oRetword.Y1Char = clscncvalue.Y1Char;
            oRetword.X2Char = clscncvalue.X2Char;
            oRetword.Y2Char = clscncvalue.Y2Char;
            oRetword.Confidence = clscncvalue.Confidence;
            oRetword.LineNo = clscncvalue.LineNo;
            oRetword.PageNo = iCurpgno;
            oRetword.Left = clscncvalue.Left;
            oRetword.Right = clscncvalue.Right;
            oRetword.Top = clscncvalue.Top;
            oRetword.Bottom = clscncvalue.Bottom;
            oRetword.strWord = strval.Trim();
            oRetword.ConfString = clscncvalue.ConfString;
            return oRetword;
        }
        #endregion

        #region Supporting Functions
        private clsCnCWord GetDefaultWord(int iCurpgno, string strval)
        {
            clsCnCWord oRetword = new clsCnCWord();
            oRetword.X1Char = "5";
            oRetword.Y1Char = "5";
            oRetword.X2Char = "10";
            oRetword.Y2Char = "10";
            oRetword.Confidence = 90;
            oRetword.LineNo = 1;
            oRetword.PageNo = iCurpgno;
            oRetword.Left = 10;
            oRetword.Right = 20;
            oRetword.Top = 10;
            oRetword.Bottom = 20;
            oRetword.strWord = strval.ToString();
            oRetword.ConfString = "9".PadLeft(strval.Length, '9');
            return oRetword;
        }

        private clsCnCWord GetDefaultWordTerms(int iCurpgno, string strval)
        {
            clsCnCWord oRetword = new clsCnCWord();
            oRetword.X1Char = "5";
            oRetword.Y1Char = "5";
            oRetword.X2Char = "10";
            oRetword.Y2Char = "10";
            oRetword.Confidence = 90;
            oRetword.LineNo = 1;
            oRetword.PageNo = iCurpgno;
            oRetword.Left = 10;
            oRetword.Right = 20;
            oRetword.Top = 10;
            oRetword.Bottom = 20;
            oRetword.strWord = strval.ToString();
            oRetword.ConfString = "9".PadLeft(strval.Length, '9');
            oRetword.Flag = "1";
            oRetword.Remarks = "3F";
            return oRetword;
        }

        private clsCnCWord GetDefaultWordTermsRed(int iCurpgno, string strval)
        {
            clsCnCWord oRetword = new clsCnCWord();
            oRetword.X1Char = "5";
            oRetword.Y1Char = "5";
            oRetword.X2Char = "10";
            oRetword.Y2Char = "10";
            oRetword.Confidence = 90;
            oRetword.LineNo = 1;
            oRetword.PageNo = iCurpgno;
            oRetword.Left = 10;
            oRetword.Right = 20;
            oRetword.Top = 10;
            oRetword.Bottom = 20;
            oRetword.strWord = strval.ToString();
            oRetword.ConfString = "9".PadLeft(strval.Length, '9');
            oRetword.Flag = "0";
            oRetword.Remarks = "1F";
            return oRetword;
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
                if (!sLeftWords.Contains(oBoundingWords.LeftWord.strWord))
                {
                    sLeftWords = oBoundingWords.LeftWord.strWord + " " + sLeftWords;
                }
                // sLeftWords = sLeftWords & " " & oBoundingWords.LeftWord.strWord

                while (iCount < 8)
                {
                    iCount += 1;
                    GetLeftWords(metaData, oBoundingWords.LeftWord, oBoundingWords.LeftWord.PageNo, iCount, ref sLeftWords);
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Terms : GetLeftWords" + ex.Message);
            }

            return sLeftWords;
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
            catch (Exception)
            {
                //MessageBox.Show("Terms : GetBottomWords" + ex.Message);

            }

            return sTopWords;
        }
        private bool FindKeywords(string FoundWords)
        {
            if (string.IsNullOrEmpty(FoundWords.Trim()))
                return false;
            var sKeyarray = new string[] { "Freight Charge Terms:", "Freight Charge Terms :", "Delivery Terms:", "Freight Charge", "Freight Terms:", "Freight Terms;", "TERMS", "/Freight Charge Terms:" };
            for (int iLoop = 0, loopTo = sKeyarray.Length - 1; iLoop <= loopTo; iLoop++)
            {
                if (FoundWords.Trim().ToUpper().Contains(sKeyarray[iLoop].Trim().ToUpper()))
                {
                    return true;
                }
            }

            return false;
        }
        private RetStructF3 Terms_SearchKeyword(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            List<int> ConfLevel = new List<int>();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            List<int> lstLineNo = new List<int>();
            clsCnCLine currentline = new clsCnCLine();
            clsCnCWord objWord = new clsCnCWord();
            // Check the  argument and split it to get the ROI 
            int MX1 = Convert.ToInt32(strArg.Split('~')[0]);
            int MY1 = Convert.ToInt32(strArg.Split('~')[1]);
            int MX2 = Convert.ToInt32(strArg.Split('~')[2]);
            int MY2 = Convert.ToInt32(strArg.Split('~')[3]);

            DataTable wordsDictionary = ObjMetaData.DocDictionary;
            string KeywordString = "'PREPAID','COLLECT','3rd','PARTY','Prepaid','Collect','Party','|Prepaid:','Third-party'";
            string[] arrPrepaid = { "PREPAID", "Prepaid", "prepaid", "|Prepaid:" };
            string[] arrCollect = { "COLLECT", "Collect", "collect" };
            string[] arrThirdParty = { "PARTY", "party", "Third-party" };
            string[] arrCOD = { "COD", "COD:", "C.O.D." };
            string patternX = string.Format(@"\b{0}\b", "X");
            string patternY = string.Format(@"\b{0}\b", "Y");
            try
            {
                // DataRow[] dataRows = wordsDictionary.Select("([word_Otext] IN (" + KeywordString + ") OR [word_Ntext] IN (" + KeywordString + ")) AND [Word_Band] IN ('B1','B2','B3','') AND [Page_no] = " + intCurrPageNumber + " AND [X1] >= " + MX1 + " AND [Y1]>=" + MY1 + " AND [X2]<=" + MX2 + " AND [Y2]<=" + MY2 + "");
                DataRow[] dataRows = wordsDictionary.Select("([word_Otext] IN (" + KeywordString + ") OR [word_Ntext] IN (" + KeywordString + ")) AND (([Word_Band] IN ('B1','B2','B3')) OR ([Word_Band] is null )) AND [Page_no] = " + intCurrPageNumber + " AND [X1] >= " + MX1 + " AND [Y1]>=" + MY1 + " AND [X2]<=" + MX2 + " AND [Y2]<=" + MY2 + "");
                if (dataRows != null && dataRows.Length > 0)
                {
                    for (int rowIndex = 0; rowIndex < dataRows.Length; rowIndex++)
                    {
                        int LineNo = Convert.ToInt32(dataRows[rowIndex]["Line_no"]);
                        lstLineNo.Add(LineNo);
                    }
                    var distinctLines = lstLineNo.Select(x => x).Distinct().OrderBy(x => x);
                    // STEP-1(B)_1 :- check where all there words - Prepaid, collect, 3rd party are present on the same line then check it for the symbol "X"
                    foreach (var item in distinctLines)
                    {

                        currentline = ObjMetaData.Page[intCurrPageNumber].Line[item];
                        List<clsCnCWord> obj_clsCncWord_Temp = new List<clsCnCWord>();
                        // Check the each word in the line 
                        for (int i = 1; i <= currentline.WordCount; i++)
                        {
                            if (Module1.CheckValidWord(currentline.Word[i].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrPrepaid) == true)
                            {
                                if (i + 1 <= currentline.WordCount)
                                {

                                    //if (currentline.Word[i + 1].strWord.ToString().ToUpper().Trim().Contains("X") || currentline.Word[i + 1].strWord.ToString().ToUpper().Trim().Contains("Y"))
                                    if (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(currentline.Word[i + 1].strWord.ToString().ToUpper().Trim()), patternX) || Regex.IsMatch(Module1.func_RemoveSpecialCharacter(currentline.Word[i + 1].strWord.ToString().ToUpper().Trim()), patternY))
                                    {
                                        #region Dont consider COD amount 
                                        // Extend the ROI and get the all words then check all three label are present or not 
                                        clsCnCWord[] dataWords = Module1.objClsCNCForROI.GetDataForROI(ObjMetaData.Page, intCurrPageNumber, currentline.Word[i].Left - 300, currentline.Word[i].Top - 300, currentline.Word[i].Right + 300, currentline.Word[i].Bottom + 300);
                                        bool isCODFound = false;
                                        for (int index = 0; index < dataWords.Count(); index++)
                                        {
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrCOD) == true)
                                            {
                                                isCODFound = true;
                                            }
                                        }
                                        #endregion

                                        clsCnCWord termsWordLabel = new clsCnCWord();
                                        termsWordLabel = currentline.Word[i];
                                        //if (termsWordLabel is object && termsWordLabel.strWord is object && !string.IsNullOrEmpty(termsWordLabel.strWord))
                                        #region "Wrong Suspect"
                                        // Start -- Added on 7-July-2021 => if X/Y Pattern present on left side of the keyword then dont consider Trems_OBL field to avoid false positive case  eg:-  ( ) COLLECT (X) PREPAID ( )THIRD PARTY
                                        bool isWrongSuspect = false;
                                        if (i - 1 <= currentline.WordCount && i - 1 >= 1)
                                        {
                                            if (currentline.Word[i - 1].strWord.ToString() == "(" || currentline.Word[i - 1].strWord.ToString() == ")" || currentline.Word[i - 1].strWord.ToString() == "()")
                                            {
                                                isWrongSuspect = true;
                                            }
                                        }
                                        else if (isWrongSuspect == false && i - 2 <= currentline.WordCount && i - 2 >= 1)
                                        {
                                            if (currentline.Word[i - 2].strWord.ToString() == "(" || currentline.Word[i - 2].strWord.ToString() == ")" || currentline.Word[i - 2].strWord.ToString() == "()")
                                            {
                                                isWrongSuspect = true;
                                            }
                                        }
                                        else if (isWrongSuspect == false && i - 3 <= currentline.WordCount && i - 3 >= 1)
                                        {
                                            if (currentline.Word[i - 3].strWord.ToString() == "(" || currentline.Word[i - 3].strWord.ToString() == ")" || currentline.Word[i - 3].strWord.ToString() == "()")
                                            {
                                                isWrongSuspect = true;
                                            }
                                        }
                                        // End
                                        #endregion
                                        if (termsWordLabel is object && termsWordLabel.strWord is object && !string.IsNullOrEmpty(termsWordLabel.strWord) && isCODFound == false && isWrongSuspect == false)
                                        {
                                            termsWordLabel.strWord = "PPD";
                                            termsWordLabel.Flag = "0";
                                            termsWordLabel.Remarks = "1F";
                                            PossibleWords.Add(termsWordLabel);
                                        }
                                    }
                                }
                            }
                            if (Module1.CheckValidWord(currentline.Word[i].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrCollect) == true)
                            {

                                //bool contains1 = Regex.IsMatch(dBValue, pattern1);
                                if (i + 1 <= currentline.WordCount)
                                {
                                    //if (currentline.Word[i + 1].strWord.ToString().ToUpper().Trim().Contains("X") || currentline.Word[i + 1].strWord.ToString().ToUpper().Trim().Contains("Y"))
                                    if (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(currentline.Word[i + 1].strWord.ToString().ToUpper().Trim()), patternX) || Regex.IsMatch(Module1.func_RemoveSpecialCharacter(currentline.Word[i + 1].strWord.ToString().ToUpper().Trim()), patternY))
                                    {
                                        #region Dont consider COD amount
                                        // Extend the ROI and get the all words then check all three label are present or not 
                                        clsCnCWord[] dataWords = Module1.objClsCNCForROI.GetDataForROI(ObjMetaData.Page, intCurrPageNumber, currentline.Word[i].Left - 300, currentline.Word[i].Top - 300, currentline.Word[i].Right + 300, currentline.Word[i].Bottom + 300);
                                        bool isCODFound = false;
                                        for (int index = 0; index < dataWords.Count(); index++)
                                        {
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrCOD) == true)
                                            {
                                                isCODFound = true;
                                            }
                                        }
                                        #endregion
                                        clsCnCWord termsWordLabel = new clsCnCWord();
                                        termsWordLabel = currentline.Word[i];
                                        //if (termsWordLabel is object && termsWordLabel.strWord is object && !string.IsNullOrEmpty(termsWordLabel.strWord))
                                        #region "Wrong Suspect"
                                        // Start -- Added on 7-July-2021 => if X/Y Pattern present on left side of the keyword then dont consider Trems_OBL field to avoid false positive case  eg:-  ( ) COLLECT (X) PREPAID ( )THIRD PARTY
                                        bool isWrongSuspect = false;
                                        if (i - 1 <= currentline.WordCount && i - 1 >= 1)
                                        {
                                            if (currentline.Word[i - 1].strWord.ToString() == "(" || currentline.Word[i - 1].strWord.ToString() == ")" || currentline.Word[i - 1].strWord.ToString() == "()")
                                            {
                                                isWrongSuspect = true;
                                            }
                                        }
                                        else if (isWrongSuspect == false && i - 2 <= currentline.WordCount && i - 2 >= 1)
                                        {
                                            if (currentline.Word[i - 2].strWord.ToString() == "(" || currentline.Word[i - 2].strWord.ToString() == ")" || currentline.Word[i - 2].strWord.ToString() == "()")
                                            {
                                                isWrongSuspect = true;
                                            }
                                        }
                                        else if (isWrongSuspect == false && i - 3 <= currentline.WordCount && i - 3 >= 1)
                                        {
                                            if (currentline.Word[i - 3].strWord.ToString() == "(" || currentline.Word[i - 3].strWord.ToString() == ")" || currentline.Word[i - 3].strWord.ToString() == "()")
                                            {
                                                isWrongSuspect = true;
                                            }
                                        }
                                        // End
                                        #endregion
                                        if (termsWordLabel is object && termsWordLabel.strWord is object && !string.IsNullOrEmpty(termsWordLabel.strWord) && isCODFound == false && isWrongSuspect == false)
                                        {
                                            termsWordLabel.strWord = "COL";
                                            termsWordLabel.Flag = "0";
                                            termsWordLabel.Remarks = "1F";
                                            PossibleWords.Add(termsWordLabel);
                                        }
                                    }
                                }
                            }
                            if (Module1.CheckValidWord(currentline.Word[i].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrThirdParty) == true)
                            {
                                if (i + 1 <= currentline.WordCount)
                                {
                                    //if (currentline.Word[i + 1].strWord.ToString().ToUpper().Trim().Contains("X") || currentline.Word[i + 1].strWord.ToString().ToUpper().Trim().Contains("Y"))
                                    if (Regex.IsMatch(Module1.func_RemoveSpecialCharacter(currentline.Word[i + 1].strWord.ToString().ToUpper().Trim()), patternX) || Regex.IsMatch(Module1.func_RemoveSpecialCharacter(currentline.Word[i + 1].strWord.ToString().ToUpper().Trim()), patternY))
                                    {
                                        #region Dont consider COD amount
                                        // Extend the ROI and get the all words then check all three label are present or not 
                                        clsCnCWord[] dataWords = Module1.objClsCNCForROI.GetDataForROI(ObjMetaData.Page, intCurrPageNumber, currentline.Word[i].Left - 300, currentline.Word[i].Top - 300, currentline.Word[i].Right + 300, currentline.Word[i].Bottom + 300);
                                        bool isCODFound = false;
                                        for (int index = 0; index < dataWords.Count(); index++)
                                        {
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrCOD) == true)
                                            {
                                                isCODFound = true;
                                            }
                                        }
                                        #endregion
                                        clsCnCWord obj_SingleWords = new clsCnCWord();
                                        clsCnCWord termsWordLabel = new clsCnCWord();
                                        clsCnCWord termsWordLabelPrevious = new clsCnCWord();
                                        termsWordLabel = currentline.Word[i];

                                        //if (termsWordLabel is object && termsWordLabel.strWord is object && !string.IsNullOrEmpty(termsWordLabel.strWord))
                                        #region "Wrong Suspect"
                                        // Start -- Added on 7-July-2021 => if X/Y Pattern present on left side of the keyword then dont consider Trems_OBL field to avoid false positive case  eg:-  ( ) COLLECT (X) PREPAID ( )THIRD PARTY
                                        bool isWrongSuspect = false;
                                        if (i - 1 <= currentline.WordCount && i - 1 >= 1)
                                        {
                                            if (currentline.Word[i - 1].strWord.ToString() == "(" || currentline.Word[i - 1].strWord.ToString() == ")" || currentline.Word[i - 1].strWord.ToString() == "()")
                                            {
                                                isWrongSuspect = true;
                                            }
                                        }
                                        else if (isWrongSuspect == false && i - 2 <= currentline.WordCount && i - 2 >= 1)
                                        {
                                            if (currentline.Word[i - 2].strWord.ToString() == "(" || currentline.Word[i - 2].strWord.ToString() == ")" || currentline.Word[i - 2].strWord.ToString() == "()")
                                            {
                                                isWrongSuspect = true;
                                            }
                                        }
                                        else if (isWrongSuspect == false && i - 3 <= currentline.WordCount && i - 3 >= 1)
                                        {
                                            if (currentline.Word[i - 3].strWord.ToString() == "(" || currentline.Word[i - 3].strWord.ToString() == ")" || currentline.Word[i - 3].strWord.ToString() == "()")
                                            {
                                                isWrongSuspect = true;
                                            }
                                        }
                                        // End
                                        #endregion
                                        if (termsWordLabel is object && termsWordLabel.strWord is object && !string.IsNullOrEmpty(termsWordLabel.strWord) && isCODFound == false && isWrongSuspect == false)
                                        {
                                            if (i - 1 <= currentline.WordCount)
                                            {
                                                termsWordLabelPrevious = currentline.Word[i - 1];
                                            }
                                            if (termsWordLabelPrevious is object && termsWordLabelPrevious.strWord is object && !string.IsNullOrEmpty(termsWordLabelPrevious.strWord))
                                            {
                                                if (termsWordLabelPrevious.strWord.Contains("3rd") || termsWordLabelPrevious.strWord.Contains("Third") || termsWordLabelPrevious.strWord.Contains("third"))
                                                {

                                                    obj_clsCncWord_Temp.Add(termsWordLabelPrevious);
                                                    obj_clsCncWord_Temp.Add(termsWordLabel);
                                                    ClsCNC objCNC = new ClsCNC();
                                                    obj_SingleWords = objCNC.MergeWords(obj_clsCncWord_Temp.ToArray());

                                                    obj_SingleWords.strWord = "PPD";
                                                    termsWordLabel.Flag = "0";
                                                    termsWordLabel.Remarks = "1F";
                                                    PossibleWords.Add(obj_SingleWords);
                                                }
                                            }
                                            //obj_SingleWords.strWord = "PPD";
                                            //termsWordLabel.Flag = "0";
                                            //termsWordLabel.Remarks = "1F";
                                            //PossibleWords.Add(obj_SingleWords);
                                        }
                                    }
                                }
                            }
                        }

                    }
                    // STEP-1(B)_2 :-  check the left keyword and top keyword 
                    if (PossibleWords.Count == 0)
                    {
                        foreach (var item in distinctLines)
                        {
                            clsCnCLine currentline1 = new clsCnCLine();
                            List<clsCnCWord> obj_clsCncWord_Temp = new List<clsCnCWord>();
                            currentline1 = ObjMetaData.Page[intCurrPageNumber].Line[item];
                            for (int i = 1; i <= currentline1.WordCount; i++)
                            {
                                // Prepaid
                                if (Module1.CheckValidWord(currentline1.Word[i].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrPrepaid) == true)
                                {
                                    string sLeft = string.Empty;
                                    string sLeftWords = string.Empty;
                                    sLeftWords = GetLeftWords(ObjMetaData, currentline1.Word[i], intCurrPageNumber, i, ref sLeft);
                                    string sTopWords = string.Empty;
                                    sTopWords = GetTopWords(ObjMetaData, currentline1.Word[i], intCurrPageNumber);
                                    if (FindKeywords(sLeftWords.Trim()) || FindKeywords(sTopWords.Trim()))
                                    {
                                        clsCnCWord termsWordLabel = new clsCnCWord();
                                        termsWordLabel = currentline1.Word[i];
                                        // Extend the ROI and get the all words then check all three label are present or not 
                                        clsCnCWord[] dataWords = Module1.objClsCNCForROI.GetDataForROI(ObjMetaData.Page, intCurrPageNumber, currentline1.Word[i].Left - 600, currentline1.Word[i].Top - 600, currentline1.Word[i].Right + 600, currentline1.Word[i].Bottom + 600);
                                        bool isCollect = false;
                                        bool isThirdParty = false;
                                        bool isCODFound = false;
                                        for (int index = 0; index < dataWords.Count(); index++)
                                        {
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrCollect) == true)
                                            {
                                                isCollect = true;
                                                //int wno = Convert.ToInt32(dataWords[index].WordNumber);
                                                //if (ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].WordCount >= wno + 1)
                                                //{
                                                //    string sword = ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].Word[wno + 1].strWord;
                                                //    if (Regex.IsMatch(sword.ToUpper().Trim(), patternX) || Regex.IsMatch(sword.ToUpper().Trim(), patternY))
                                                //    {

                                                //    }
                                                //}
                                            }
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrThirdParty) == true)
                                            {
                                                isThirdParty = true;
                                                //int wno = Convert.ToInt32(dataWords[index].WordNumber);
                                                //if (ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].WordCount >= wno + 1)
                                                //{
                                                //    string sword = ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].Word[wno + 1].strWord;
                                                //    if (Regex.IsMatch(sword.ToUpper().Trim(), patternX) || Regex.IsMatch(sword.ToUpper().Trim(), patternY))
                                                //    {

                                                //    }
                                                //}
                                            }
                                            // check COD Amount
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrCOD) == true)
                                            {
                                                isCODFound = true;
                                            }
                                        }
                                        // check only terms with lable Ex => TERMS PREPAID
                                        if (termsWordLabel is object && termsWordLabel.strWord is object && !string.IsNullOrEmpty(termsWordLabel.strWord) && isCollect == false && isThirdParty == false && isCODFound == false)
                                        {
                                            termsWordLabel.strWord = "PPD";
                                            termsWordLabel.Flag = "0";
                                            termsWordLabel.Remarks = "1F";
                                            PossibleWords.Add(termsWordLabel);
                                        }
                                        // check in the current line words if it contains only one word and we found left and top keyword  
                                    }
                                }
                                // Collect
                                if (Module1.CheckValidWord(currentline1.Word[i].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrCollect) == true)
                                {
                                    string sLeft = string.Empty;
                                    string sLeftWords = string.Empty;
                                    sLeftWords = GetLeftWords(ObjMetaData, currentline1.Word[i], intCurrPageNumber, i, ref sLeft);
                                    string sTopWords = string.Empty;
                                    sTopWords = GetTopWords(ObjMetaData, currentline1.Word[i], intCurrPageNumber);
                                    if (FindKeywords(sLeftWords.Trim()) || FindKeywords(sTopWords.Trim()))
                                    {
                                        clsCnCWord termsWordLabel = new clsCnCWord();
                                        termsWordLabel = currentline1.Word[i];
                                        clsCnCWord[] dataWords = Module1.objClsCNCForROI.GetDataForROI(ObjMetaData.Page, intCurrPageNumber, currentline1.Word[i].Left - 600, currentline1.Word[i].Top - 600, currentline1.Word[i].Right + 600, currentline1.Word[i].Bottom + 600);
                                        bool isPrepaid = false;
                                        bool isThirdParty = false;
                                        bool isCODFound = false;
                                        // bool isThirdPartyFound = false;
                                        for (int index = 0; index < dataWords.Count(); index++)
                                        {
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrPrepaid) == true)
                                            {
                                                isPrepaid = true;
                                                //int wno = Convert.ToInt32(dataWords[index].WordNumber);
                                                //if (ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].WordCount >= wno + 1)
                                                //{
                                                //    string strword = ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].Word[wno + 1].strWord;
                                                //    if (Regex.IsMatch(strword.ToUpper().Trim(), patternX) || Regex.IsMatch(strword.ToUpper().Trim(), patternY))
                                                //    {
                                                //        isPrepaidFound = false;
                                                //    }
                                                //}
                                            }
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrThirdParty) == true)
                                            {
                                                isThirdParty = true;
                                                //int wno = Convert.ToInt32(dataWords[index].WordNumber);
                                                //if (ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].WordCount >= wno + 1)
                                                //{
                                                //    string strword = ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].Word[wno + 1].strWord;
                                                //    if (Regex.IsMatch(strword.ToUpper().Trim(), patternX) || Regex.IsMatch(strword.ToUpper().Trim(), patternY))
                                                //    {
                                                //        isThirdPartyFound = false;
                                                //    }
                                                //}
                                            }
                                            // check COD Amount
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrCOD) == true)
                                            {
                                                isCODFound = true;
                                            }
                                        }
                                        // check only terms with lable Ex => TERMS COLLECT
                                        if (termsWordLabel is object && termsWordLabel.strWord is object && !string.IsNullOrEmpty(termsWordLabel.strWord) && isPrepaid == false && isThirdParty == false && isCODFound == false)
                                        {
                                            termsWordLabel.strWord = "COL";
                                            termsWordLabel.Flag = "0";
                                            termsWordLabel.Remarks = "1F";
                                            PossibleWords.Add(termsWordLabel);
                                        }
                                        //else
                                        //{
                                        //    //  Ex => TERMS PREPAID
                                        //    //              COLLECT
                                        //    //              Third party
                                        //    // check all three lable in the same ROI 
                                        //    if (termsWordLabel is object && termsWordLabel.strWord is object && !string.IsNullOrEmpty(termsWordLabel.strWord) && isPrepaid == true && isThirdParty == true)
                                        //    {
                                        //        // check if both are tick then dont not do anything if any one is tick then 
                                        //        termsWordLabel.Flag = "0";
                                        //        termsWordLabel.Remarks = "1F";
                                        //        PossibleWords.Add(termsWordLabel);
                                        //    }
                                        //}

                                    }
                                }
                                // 3rd party
                                if (Module1.CheckValidWord(currentline1.Word[i].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrThirdParty) == true)
                                {
                                    string sLeft = string.Empty;
                                    string sLeftWords = string.Empty;
                                    sLeftWords = GetLeftWords(ObjMetaData, currentline1.Word[i], intCurrPageNumber, i, ref sLeft);
                                    string sTopWords = string.Empty;
                                    sTopWords = GetTopWords(ObjMetaData, currentline1.Word[i], intCurrPageNumber);
                                    if (FindKeywords(sLeftWords.Trim()) || FindKeywords(sTopWords.Trim()))
                                    {
                                        clsCnCWord obj_SingleWords = new clsCnCWord();
                                        clsCnCWord termsWordLabel = new clsCnCWord();
                                        clsCnCWord termsWordLabelPrevious = new clsCnCWord();
                                        termsWordLabel = currentline1.Word[i];
                                        clsCnCWord[] dataWords = Module1.objClsCNCForROI.GetDataForROI(ObjMetaData.Page, intCurrPageNumber, currentline1.Word[i].Left - 600, currentline1.Word[i].Top - 600, currentline1.Word[i].Right + 600, currentline1.Word[i].Bottom + 600);
                                        bool isPrepaid = false;
                                        bool isCollect = false;
                                        bool isCODFound = false;
                                        for (int index = 0; index < dataWords.Count(); index++)
                                        {
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrPrepaid) == true)
                                            {
                                                isPrepaid = true;
                                                //int wno = Convert.ToInt32(dataWords[index].WordNumber);
                                                //if (ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].WordCount >= wno + 1)
                                                //{
                                                //    string strword = ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].Word[wno + 1].strWord;
                                                //    if (Regex.IsMatch(strword.ToUpper().Trim(), patternX) || Regex.IsMatch(strword.ToUpper().Trim(), patternY))
                                                //    {

                                                //    }
                                                //}
                                            }
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrCollect) == true)
                                            {
                                                isCollect = true;
                                                //int wno = Convert.ToInt32(dataWords[index].WordNumber);
                                                //if (ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].WordCount >= wno + 1)
                                                //{
                                                //    string strword = ObjMetaData.Page[intCurrPageNumber].Line[dataWords[index].LineNo].Word[wno + 1].strWord;
                                                //    if (Regex.IsMatch(strword.ToUpper().Trim(), patternX) || Regex.IsMatch(strword.ToUpper().Trim(), patternY))
                                                //    {

                                                //    }
                                                //}
                                            }
                                            // check COD Amount
                                            if (Module1.CheckValidWord(dataWords[index].strWord.ToString().ToUpper().Trim().Replace("'", "`"), arrCOD) == true)
                                            {
                                                isCODFound = true;
                                            }
                                        }

                                        // check only terms with lable Ex => TERMS Third party
                                        if (termsWordLabel is object && termsWordLabel.strWord is object && !string.IsNullOrEmpty(termsWordLabel.strWord) && isPrepaid == false && isCollect == false && isCODFound == false)
                                        {
                                            if (i - 1 <= currentline1.WordCount)
                                            {
                                                termsWordLabelPrevious = currentline1.Word[i - 1];
                                            }
                                            if (termsWordLabelPrevious is object && termsWordLabelPrevious.strWord is object && !string.IsNullOrEmpty(termsWordLabelPrevious.strWord))
                                            {
                                                if (termsWordLabelPrevious.strWord.Contains("3rd") || termsWordLabelPrevious.strWord.Contains("Third") || termsWordLabelPrevious.strWord.Contains("third"))
                                                {
                                                    obj_clsCncWord_Temp.Add(termsWordLabelPrevious);
                                                    obj_clsCncWord_Temp.Add(termsWordLabel);
                                                    ClsCNC objCNC = new ClsCNC();
                                                    obj_SingleWords = objCNC.MergeWords(obj_clsCncWord_Temp.ToArray());
                                                    obj_SingleWords.strWord = "PPD";
                                                    obj_SingleWords.Flag = "0";
                                                    obj_SingleWords.Remarks = "1F";
                                                    PossibleWords.Add(obj_SingleWords);
                                                }
                                            }
                                            //obj_SingleWords.strWord = "PPD";
                                            //obj_SingleWords.Flag = "0";
                                            //obj_SingleWords.Remarks = "1F";
                                            //PossibleWords.Add(obj_SingleWords);
                                        }
                                    }
                                }
                            }
                        }
                    }

                }
                if (PossibleWords.Count == 1)
                {
                    Module1.objRetStructF3.Status = "S";
                    ConfLevel.Add(90);
                    Module1.objRetStructF3.Words = PossibleWords;
                    Module1.objRetStructF3.ManualConfirmation = "N";
                    Module1.objRetStructF3.Flag = "Y";
                    Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
                    Module1.objRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                }

                //wordsDictionary.Dispose();
                //wordsDictionary = null;

                //return Module1.objRetStructF3;
            }
            catch (Exception ex)
            {

            }
            finally
            {
                if (wordsDictionary != null)
                {
                    wordsDictionary.Dispose();
                    wordsDictionary = null;
                }
                PossibleWords = null;
                ConfLevel = null;
                lstLineNo = null;
            }
            return Module1.objRetStructF3;
        }

        #endregion

        public class ClsOCRValues
        {
            // Shipper
            public static string ClsOCRValuesShipperName { get; set; }
            public static string ClsOCRValuesShipperAddressLine1 { get; set; }
            public static string ClsOCRValuesShipperCity { get; set; }
            public static string ClsOCRValuesShipperState { get; set; }
            public static string ClsOCRValuesShipperZip { get; set; }
            public static string ClsOCRValuesShipperTelePhoneNo { get; set; }

            // Consignee
            public static string ClsOCRValuesConsigneeName { get; set; }
            public static string ClsOCRValuesConsigneeAddressLine1 { get; set; }
            public static string ClsOCRValuesConsigneeCity { get; set; }
            public static string ClsOCRValuesConsigneeState { get; set; }
            public static string ClsOCRValuesConsigneeZip { get; set; }
            public static string ClsOCRValuesConsigneeTelePhoneNo { get; set; }

            // Bill to 3rd party

            public static string ClsOCRValuesBillToName { get; set; }
            public static string ClsOCRValuesBillToAddressLine1 { get; set; }
            public static string ClsOCRValuesBillToCity { get; set; }
            public static string ClsOCRValuesBillToState { get; set; }
            public static string ClsOCRValuesBillToZip { get; set; }
            public static string ClsOCRValuesBillToTelePhoneNo { get; set; }

            // OCR Confidence level
            // For Name
            public static string ShipperNameOCRWithHighConfidence { get; set; }
            public static string ConsigneeNameOCRWithHighConfidence { get; set; }
            public static string BillToNameOCRWithHighConfidence { get; set; }

            public static string ShipperTelephoneOCRWithHighConfidence { get; set; }
            public static string ConsigneeTelephoneOCRWithHighConfidence { get; set; }
            public static string BillToTelephoneOCRWithHighConfidence { get; set; }

            // For address => consignee

            public static string ConsigneeAddressLine1OCRWithHighConfidence { get; set; }
            public static string ConsigneeCityOCRWithHighConfidence { get; set; }
            public static string ConsigneeStateOCRWithHighConfidence { get; set; }
            public static string ConsigneeZipOCRWithHighConfidence { get; set; }

            // Terms 
            public static string ClsOCRValuesTerms { get; set; }

            // added on 7-July-2021 => tag correct consignee
            public static string ClsOCRCorrectConsigneeTagged { get; set; }

            // Added on 30-Dec-2021 => EDI document
            public static string IsEDIDocument { get; set; }
        }


    }
}