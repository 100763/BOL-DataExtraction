using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Data.Odbc;
using System.Configuration;
using System.Linq;
using System.Text.RegularExpressions;

namespace BOLSET2IQ
{
    public class clsF27
    {
        string mdbFileConnString = Application.StartupPath + "\\" + "Combine.mdb";
        string As400_ConnectionString = "";//"Driver={iSeries Access ODBC Driver}; System=SAIATEST;UID=BOTBOLOCR;PWD=saia789$;Naming=1;DBQ=* USRLIBL; Compression=1";
        public RetStructF27 F27(string strRoutineNo, string strArg, string strFldId, string currRowNo, string type)
        {
            RetStructF27 objStructF27 = new RetStructF27();
            objStructF27.Status = "F";
            clsCNCSBR ObjSbr = new clsCNCSBR();
            RetStructIQSBR050 objRet50 = new RetStructIQSBR050();
            RetStructIQSBR058 objRet58 = new RetStructIQSBR058();
            if (Module1.ActualimageName.ToLower() != Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath).ToLower())
            {
                //*********************DB2Conn***************
                //if (string.IsNullOrEmpty(iQPUBLIC.PublicComponents.DbConnectionString1))
                //{
                //    iQPUBLIC.PublicComponents.DbConnectionString1 = ConfigurationManager.ConnectionStrings["DB2Conn"].ToString();  //DB2
                //}
                //**********************************************

                #region Set empty => Global properties
                // Get all Shipper details
                clsModule.clsPickupValuesFRP001ShipperDetailsPickUp = string.Empty;
                clsModule.clsPickupValuesFRP001ShipperCode = string.Empty;
                clsModule.clsPickupValuesFRP001ShipperName = string.Empty;
                clsModule.clsPickupValuesFRP001ShipperAddress1 = string.Empty;
                clsModule.clsPickupValuesFRP001ShipperCity = string.Empty;
                clsModule.clsPickupValuesFRP001ShipperState = string.Empty;
                clsModule.clsPickupValuesFRP001ShipperZip = string.Empty;
                clsModule.clsPickupValuesFRP001ShipperTerminalID = string.Empty;
                // Get all details of consignee
                clsModule.clsPickupValuesFRP001ConsigneeCode = string.Empty;
                clsModule.clsPickupValuesFRP001ConsigneeName = string.Empty;
                clsModule.clsPickupValuesFRP001ConsigneeAddress1 = string.Empty;
                clsModule.clsPickupValuesFRP001ConsigneeCity = string.Empty;
                clsModule.clsPickupValuesFRP001ConsigneeState = string.Empty;
                clsModule.clsPickupValuesFRP001ConsigneeZip = string.Empty;
                clsModule.clsPickupValuesFRP001ConsigneeTerminalID = string.Empty;
                // Get all details of Bill To 
                clsModule.clsPickupValuesFRP001BillToCode = string.Empty;
                clsModule.clsPickupValuesFRP001BillToName = string.Empty;
                clsModule.clsPickupValuesFRP001BillToAddress1 = string.Empty;
                clsModule.clsPickupValuesFRP001BillToCity = string.Empty;
                clsModule.clsPickupValuesFRP001BillToState = string.Empty;
                clsModule.clsPickupValuesFRP001BillToZip = string.Empty;
                // RAT Bill For Shipper
                clsModule.clsPickupValuesFRP001RAT = string.Empty;
                // Get Blind shipper (BS)
                clsModule.clsPickupValuesFRP001BS = string.Empty;
                // HAT Bill for Consignee
                clsModule.clsPickupValuesFRP001HAT = string.Empty;

                // DB Values -Shipper
                // Shipper attached to BillTo
                clsModule.isCMBTPCMBTCpresentShipper = string.Empty;
                clsModule.isCMBTPCMBTCpresentValueShipper = string.Empty;
                clsModule.ShipperPrepaidCodeDetail = string.Empty;
                clsModule.ShipperCollectCodeDetail = string.Empty; // NOT IN USE
                // Shipper  RAT
                clsModule.clsMasterValuesAPR001ShipperRATCode = string.Empty;
                clsModule.clsMasterValuesAPR001ShipperRATDetails = string.Empty;
                clsModule.clsMasterValuesAPR001ShipperRATCodeFinal = string.Empty;
                clsModule.clsMasterValuesAPR001ShipperRATTelePhoneNo = string.Empty;
                // Shipper BS
                clsModule.clsPickupValuesFRP001BSMatched = string.Empty;
                clsModule.clsPickupValuesFRP001BSMatchedCode = string.Empty;
                clsModule.clsPickupValuesFRP001BSMatchedDetails = string.Empty;

                // DB Values - Consignee
                clsModule.isCMBTPCMBTCpresent = string.Empty;
                clsModule.isCMBTPCMBTCpresentValue = string.Empty;
                clsModule.ConsigneePrepaidCodeDetail = string.Empty; // NOT IN USE
                clsModule.ConsigneeCollectCodeDetail = string.Empty;
                // consignee HAT
                clsModule.clsMasterValuesAPR001ConsigneeHATCode = string.Empty;
                clsModule.clsMasterValuesAPR001ConsigneeHATDetails = string.Empty;
                clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal = string.Empty;
                clsModule.clsMasterValuesAPR001ConsigneeHATTelePhoneNo = string.Empty;

                // Final Values => 
                // shipper => 
                clsModule.clsFinalValuesShipperName = string.Empty;
                clsModule.clsFinalValuesShipperAddress1 = string.Empty;
                clsModule.clsFinalValuesShipperCity = string.Empty;
                clsModule.clsFinalValuesShipperState = string.Empty;
                clsModule.clsFinalValuesShipperZip = string.Empty;
                // Consignee
                clsModule.clsFinalValuesConsigneeName = string.Empty;
                clsModule.clsFinalValuesConsigneeAddress1 = string.Empty;
                clsModule.clsFinalValuesConsigneeCity = string.Empty;
                clsModule.clsFinalValuesConsigneeState = string.Empty;
                clsModule.clsFinalValuesConsigneeZip = string.Empty;

                // Bill To
                clsModule.clsFinalValuesBillToName = string.Empty;
                clsModule.clsFinalValuesBillToAddress1 = string.Empty;
                clsModule.clsFinalValuesBillToCity = string.Empty;
                clsModule.clsFinalValuesBillToState = string.Empty;
                clsModule.clsFinalValuesBillToZip = string.Empty;

                // Mail to attached to shipper/ consignee
                clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail = string.Empty;
                clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail = string.Empty;

                // Address1 and Address2 with pipe separated
                clsModule.clsMasterValuesAPR001ShipperDetailsWithPipe = string.Empty;
                //clsModule.clsMasterValuesAPR001ShipperRATDetailsWithPipe = string.Empty;
                //clsModule.clsMasterValuesAPR001ShipperBSDetailsWithPipe = string.Empty;

                // consignee
                clsModule.clsMasterValuesAPR001ConsigneeDetailsWithPipe = string.Empty;
                //clsModule.clsMasterValuesAPR001ConsigneeHATDetailsWithPipe = string.Empty;

                // 3rd party
                clsModule.clsMasterValuesAPR001MailToAttachedShipperWithPipe = string.Empty;
                clsModule.clsMasterValuesAPR001MailToAttachedConsigneeWithPipe = string.Empty;
                clsModule.clsMasterValuesAPR001BillToAttachedShipperWithPipe = string.Empty;
                clsModule.clsMasterValuesAPR001BillToAttachedConsigneeWithPipe = string.Empty;

                // Serach engine 
                ClsSearchEngin.ClsSearchEnginShipperDetailsWithPipe = string.Empty;
                ClsSearchEngin.ClsSearchEnginConsigneeDetailsWithPipe = string.Empty;
                ClsSearchEngin.ClsSearchEnginBillToDetailsWithPipe = string.Empty;

                // alise case
                ClsSearchEngin.ClsSearchEnginAliseShipperNameARP033 = string.Empty;
                ClsSearchEngin.ClsSearchEnginAliseConsigneeNameARP033 = string.Empty;
                ClsSearchEngin.ClsSearchEnginAliseBillToNameARP033 = string.Empty;

                clsModule.clsMasterValuesAPR001AliseShipperName = string.Empty;
                clsModule.clsMasterValuesAPR001AliseConsigneeName = string.Empty;
                clsModule.clsMasterValuesAPR001AliseBillToName = string.Empty;

                ClsSearchEngin.ClsSearchCommonStatusFlag = string.Empty;
                ClsSearchEngin.ClsSearchCommonStatusFlagShipper = string.Empty;
                ClsSearchEngin.ClsSearchCommonStatusFlagConsignee = string.Empty;
                ClsSearchEngin.ClsSearchCommonStatusFlagBillTo = string.Empty;

                // Search using Telephone
                clsModule.clsMasterValuesAPR001ShipperCodeUsingTelphone = string.Empty;
                clsModule.clsMasterValuesAPR001ConsigneeCodeUsingTelphone = string.Empty;
                clsModule.clsMasterValuesAPR001BillToCodeUsingTelphone = string.Empty;

                // Valid code 
                clsModule.isShipperValidFlag = "T";
                clsModule.isConsigneeValidFlag = "T";
                // find consignee miscellaneous code
                clsModule.isConsigneeCodeRedFlag = "F";
                clsModule.isConsigneeAllGreenFlag = "T";
                if (Module1.TableRecordssFinal != null)
                {
                    Module1.TableRecordssFinal.Dispose();
                    Module1.TableRecordssFinal = null;
                }
                // Added on 13-July-2021 
                clsModule.ShipperKMFinalCode = string.Empty;
                clsModule.ShipperKMFinalName = string.Empty;
                clsModule.ShipperKMFinalAddressLine = string.Empty;
                clsModule.ShipperKMFinalCity = string.Empty;
                clsModule.ShipperKMFinalState = string.Empty;
                clsModule.ShipperKMFinalZip = string.Empty;
                clsModule.ShipperKMFlag = "F";
                clsModule.ShipperKMOBLName = string.Empty;
                clsModule.ShipperKMOBLAddressLine = string.Empty;
                #endregion

                if (Module1.dtUSCityStateZip.Rows.Count <= 0)
                {
                    Module1.dtUSCityStateZip = Module1.GetUSCityStateZipCode();
                }
                Module1.ActualimageName = Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath).ToLower();
                Module1.imageName = Path.GetFileNameWithoutExtension(iQPUBLIC.PublicComponents.PrimaryImagePath);
                Module1.imageName = Module1.imageName.Split('_')[Module1.imageName.Split('_').Length - 1];
                if (!string.IsNullOrEmpty(Module1.imageName))
                {
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET2_PickupShipperDetails") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET2_PickupShipperDetails");
                    }
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET3_TotalWeight") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET3_TotalWeight");
                    }
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET3_TotalPieces") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET3_TotalPieces");
                    }

                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET2_ShipperValidFlag"))
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET2_ShipperValidFlag");
                    }
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET2_ConsigneeMiscellaneousFlag"))
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET2_ConsigneeMiscellaneousFlag");
                    }


                    string queryStringShipper = "select FHSWGT, FHTOTP, FHSCD as PASCD,FHSNM as PASNM,FHSA1 as PASA1,FHSA2 as PASA2,FHSCT as PASCT,FHSST as PASST,FHSZIP as PASZIP,FHOT as ShipperTerminalID, FHCCD as ConsigneeCode,FHCNM as ConsigneeName, FHCA1 as ConsigneeAddress1, FHCA2 as ConsigneeAddress2, FHCCT as ConsigneeCity, FHCST as ConsigneeState, FHCZIP as ConsigneeZip, FHDT as ConsigneeTerminalID, FHBTC as BillToCode, FHBNM as BillToName, FHBA1 as BillToAddress1, FHBA2 as BillToAddress2, FHBCT as BillToCity, FHBST as BillToState, FHBST as BillToCity, FHBZIP as BillToZipCode, case when FHSA1 like '%EXIT SAIA%' or FHSA1 like '%RAT EXIT%' or FHSA1 like '%%SAIA%'  then 'Y' else 'N' end as RATBill, case when FHCA1 like '%%SAIA%' or FHCA1 like '%% SAIA%'  then 'Y' else 'N' end as HATBill from FRP001 where FHPRO='" + Module1.imageName + "'";
                    DataTable dtPickup_Shipper_Details = GetDataTableByText(As400_ConnectionString, queryStringShipper);
                    //DataTable dtPickup_Shipper_Details = GetDataTableByText(queryStringShipper);
                    if (dtPickup_Shipper_Details != null && dtPickup_Shipper_Details.Rows.Count > 0)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET2_PickupShipperDetails", dtPickup_Shipper_Details);
                        //********************total_Weight*************

                        string total_Weight = "";
                        if (!DBNull.Value.Equals(dtPickup_Shipper_Details.Rows[0]["FHSWGT"]))
                        {
                            total_Weight = dtPickup_Shipper_Details.Rows[0]["FHSWGT"].ToString().Trim();
                        }
                        iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET3_TotalWeight", total_Weight);

                        //********************total_Pieces**********************

                        string total_Pieces = "";
                        if (!DBNull.Value.Equals(dtPickup_Shipper_Details.Rows[0]["FHTOTP"]))
                        {
                            total_Pieces = dtPickup_Shipper_Details.Rows[0]["FHTOTP"].ToString().Trim();
                        }
                        iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET3_TotalPieces", total_Pieces);
                        // Get all Shipper details
                        //shipperCodePickup = dtPickup_Shipper_Details.Rows[0]["PASCD"].ToString();
                        if (dtPickup_Shipper_Details.Rows[0]["PASA1"].ToString().Trim() == dtPickup_Shipper_Details.Rows[0]["PASA2"].ToString().Trim())
                        {
                            clsModule.clsPickupValuesFRP001ShipperDetailsPickUp = dtPickup_Shipper_Details.Rows[0]["PASNM"].ToString().Trim() + "~" + dtPickup_Shipper_Details.Rows[0]["PASA1"].ToString().Trim() + "~" + dtPickup_Shipper_Details.Rows[0]["PASCT"].ToString().Trim() + "~" + dtPickup_Shipper_Details.Rows[0]["PASST"].ToString().Trim() + "~" + dtPickup_Shipper_Details.Rows[0]["PASZIP"].ToString().Trim();
                        }
                        else
                        {
                            clsModule.clsPickupValuesFRP001ShipperDetailsPickUp = dtPickup_Shipper_Details.Rows[0]["PASNM"].ToString().Trim() + "~" + dtPickup_Shipper_Details.Rows[0]["PASA1"].ToString().Trim() + " " + dtPickup_Shipper_Details.Rows[0]["PASA2"].ToString().Trim() + "~" + dtPickup_Shipper_Details.Rows[0]["PASCT"].ToString().Trim() + "~" + dtPickup_Shipper_Details.Rows[0]["PASST"].ToString().Trim() + "~" + dtPickup_Shipper_Details.Rows[0]["PASZIP"].ToString().Trim();
                        }
                        clsModule.clsPickupValuesFRP001ShipperCode = dtPickup_Shipper_Details.Rows[0]["PASCD"].ToString().Trim();
                        clsModule.clsPickupValuesFRP001ShipperName = dtPickup_Shipper_Details.Rows[0]["PASNM"].ToString().Trim();
                        if (dtPickup_Shipper_Details.Rows[0]["PASA1"].ToString().Trim() == dtPickup_Shipper_Details.Rows[0]["PASA2"].ToString().Trim())
                        {
                            clsModule.clsPickupValuesFRP001ShipperAddress1 = dtPickup_Shipper_Details.Rows[0]["PASA1"].ToString().Trim();
                        }
                        else
                        {
                            clsModule.clsPickupValuesFRP001ShipperAddress1 = dtPickup_Shipper_Details.Rows[0]["PASA1"].ToString().Trim() + " " + dtPickup_Shipper_Details.Rows[0]["PASA2"].ToString().Trim();
                        }
                        clsModule.clsPickupValuesFRP001ShipperCity = dtPickup_Shipper_Details.Rows[0]["PASCT"].ToString().Trim();
                        clsModule.clsPickupValuesFRP001ShipperState = dtPickup_Shipper_Details.Rows[0]["PASST"].ToString().Trim();
                        clsModule.clsPickupValuesFRP001ShipperZip = dtPickup_Shipper_Details.Rows[0]["PASZIP"].ToString().Trim();
                        clsModule.clsPickupValuesFRP001ShipperTerminalID = dtPickup_Shipper_Details.Rows[0]["ShipperTerminalID"].ToString().Trim();
                        // Get all details of consignee
                        clsModule.clsPickupValuesFRP001ConsigneeCode = dtPickup_Shipper_Details.Rows[0]["ConsigneeCode"].ToString().Trim();
                        clsModule.clsPickupValuesFRP001ConsigneeName = dtPickup_Shipper_Details.Rows[0]["ConsigneeName"].ToString().Trim();
                        if (dtPickup_Shipper_Details.Rows[0]["ConsigneeAddress1"].ToString().Trim() == dtPickup_Shipper_Details.Rows[0]["ConsigneeAddress2"].ToString().Trim())
                        {
                            clsModule.clsPickupValuesFRP001ConsigneeAddress1 = dtPickup_Shipper_Details.Rows[0]["ConsigneeAddress1"].ToString().Trim();
                        }
                        else
                        {
                            clsModule.clsPickupValuesFRP001ConsigneeAddress1 = dtPickup_Shipper_Details.Rows[0]["ConsigneeAddress1"].ToString().Trim() + " " + dtPickup_Shipper_Details.Rows[0]["ConsigneeAddress2"].ToString().Trim();
                        }
                        clsModule.clsPickupValuesFRP001ConsigneeCity = dtPickup_Shipper_Details.Rows[0]["ConsigneeCity"].ToString().Trim();
                        clsModule.clsPickupValuesFRP001ConsigneeState = dtPickup_Shipper_Details.Rows[0]["ConsigneeState"].ToString().Trim();
                        clsModule.clsPickupValuesFRP001ConsigneeZip = dtPickup_Shipper_Details.Rows[0]["ConsigneeZip"].ToString().Trim();
                        clsModule.clsPickupValuesFRP001ConsigneeTerminalID = dtPickup_Shipper_Details.Rows[0]["ConsigneeTerminalID"].ToString().Trim();

                        // Get all details of Bill To 
                        clsModule.clsPickupValuesFRP001BillToCode = dtPickup_Shipper_Details.Rows[0]["BillToCode"].ToString().Trim();
                        clsModule.clsPickupValuesFRP001BillToName = dtPickup_Shipper_Details.Rows[0]["BillToName"].ToString().Trim();
                        if (dtPickup_Shipper_Details.Rows[0]["BillToAddress1"].ToString().Trim() == dtPickup_Shipper_Details.Rows[0]["BillToAddress2"].ToString().Trim())
                        {
                            clsModule.clsPickupValuesFRP001BillToAddress1 = dtPickup_Shipper_Details.Rows[0]["BillToAddress1"].ToString().Trim();
                        }
                        else
                        {
                            clsModule.clsPickupValuesFRP001BillToAddress1 = dtPickup_Shipper_Details.Rows[0]["BillToAddress1"].ToString().Trim() + " " + dtPickup_Shipper_Details.Rows[0]["BillToAddress2"].ToString().Trim();
                        }
                        clsModule.clsPickupValuesFRP001BillToCity = dtPickup_Shipper_Details.Rows[0]["BillToCity"].ToString().Trim();
                        clsModule.clsPickupValuesFRP001BillToState = dtPickup_Shipper_Details.Rows[0]["BillToState"].ToString().Trim();
                        clsModule.clsPickupValuesFRP001BillToZip = dtPickup_Shipper_Details.Rows[0]["BillToZipCode"].ToString().Trim();
                        // RAT Bill For Shipper
                        clsModule.clsPickupValuesFRP001RAT = dtPickup_Shipper_Details.Rows[0]["RATBill"].ToString().Trim();
                        // Get Blind shipper (BS)
                        if (clsModule.clsPickupValuesFRP001ShipperName.EndsWith("(BS)") || clsModule.clsPickupValuesFRP001ShipperName.EndsWith("BS"))
                        {
                            clsModule.clsPickupValuesFRP001BS = "Y";
                        }
                        else
                        {
                            clsModule.clsPickupValuesFRP001BS = "N";
                        }
                        //} 
                        // HAT Bill for Consignee
                        clsModule.clsPickupValuesFRP001HAT = dtPickup_Shipper_Details.Rows[0]["HATBill"].ToString().Trim();
                        if (dtPickup_Shipper_Details != null)
                        {
                            dtPickup_Shipper_Details.Dispose();
                            dtPickup_Shipper_Details = null;
                        }

                    }
                }
            }
            try
            {
                switch (strRoutineNo)
                {
                    case "C01": //  Shipper Code# => Pickup 
                        if (!string.IsNullOrEmpty(clsModule.clsPickupValuesFRP001ShipperCode))
                        {
                            clsModule.clsPickupValuesFRP001ShipperCode = clsModule.clsPickupValuesFRP001ShipperCode.PadLeft(7, '0');
                            objStructF27.ContentsOfField = clsModule.clsPickupValuesFRP001ShipperCode;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;
                    case "C02": // Shipper Details (~Sep) => Pickup => Name~Add~City~State~Zip~TelNo
                        if (!string.IsNullOrEmpty(clsModule.clsPickupValuesFRP001ShipperDetailsPickUp))
                        {
                            objStructF27.ContentsOfField = clsModule.clsPickupValuesFRP001ShipperDetailsPickUp;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;
                    case "C09": // Shipper Master KeyField =>  Code_Pickup/TelNo_OCR
                        if (clsModule.clsPickupValuesFRP001RAT == "Y" || (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "@@@@@"))
                        {
                            string strshipperCodePickup = clsModule.clsPickupValuesFRP001ShipperCode;
                            string strRAT = clsModule.clsPickupValuesFRP001RAT;
                            if (!string.IsNullOrEmpty(strshipperCodePickup))
                            {
                                string strShipperCodeMasterOrTelNoOCR = string.Empty;
                                string strShipperDetailsMaster = string.Empty;
                                string strflag = "S";
                                GetShipperDetailsMaster(strshipperCodePickup, strflag, strRAT, ref strShipperCodeMasterOrTelNoOCR, ref strShipperDetailsMaster);
                                string ShipperCodeMasterOrTelOCR = strShipperCodeMasterOrTelNoOCR;
                                if (!string.IsNullOrEmpty(ShipperCodeMasterOrTelOCR))
                                {
                                    objStructF27.ContentsOfField = ShipperCodeMasterOrTelOCR;
                                }
                                else
                                {
                                    objStructF27.ContentsOfField = "@@@@@";
                                }
                            }
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "NA";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;
                    case "C10": // Shipper Details (~Sep) =>  Master

                        if (clsModule.clsPickupValuesFRP001RAT == "Y" || (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "@@@@@"))
                        {
                            string strshipperCodePickupDetail = clsModule.clsPickupValuesFRP001ShipperCode;
                            string strRAT = clsModule.clsPickupValuesFRP001RAT;
                            if (!string.IsNullOrEmpty(strshipperCodePickupDetail))
                            {
                                string strShipperCodeMasterOrTelNoOCR = string.Empty;
                                string strShipperDetailsMaster = string.Empty;
                                string strflag = "S";
                                GetShipperDetailsMaster(strshipperCodePickupDetail, strflag, strRAT, ref strShipperCodeMasterOrTelNoOCR, ref strShipperDetailsMaster);
                                string ShipperDetailMaster = strShipperDetailsMaster;
                                if (!string.IsNullOrEmpty(ShipperDetailMaster))
                                {
                                    objStructF27.ContentsOfField = ShipperDetailMaster;
                                }
                                else
                                {
                                    objStructF27.ContentsOfField = "@@@@@";
                                }
                            }
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "NA";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;
                    case "C11": // Shipper Details OCR_Matched_MasterDB_Key => Yes/No
                        string strLogic = string.Empty;
                        string IsMatchNameShipper = string.Empty;
                        // check for RAT
                        if ((clsModule.clsPickupValuesFRP001RAT == "Y"))
                        {
                            if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATDetails))
                            {
                                strLogic = "YES";
                            }
                            else
                            {
                                strLogic = "NO";
                            }
                        }
                        // Check for BS
                        else if (clsModule.clsPickupValuesFRP001BS == "Y")
                        {
                            if (clsModule.clsPickupValuesFRP001BSMatched == "Y")
                            {
                                strLogic = "YES";
                            }
                            else
                            {
                                strLogic = "NO";
                            }
                        }
                        else // Check for normal case
                        {
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "@@@@@")
                            {
                                string shipperDetilsMasterCell = clsModule.clsMasterValuesAPR001ShipperDetails;
                                if (!string.IsNullOrEmpty(shipperDetilsMasterCell) && shipperDetilsMasterCell != "@@@@@")
                                {
                                    string[] arrshipperDetilsMasterCell = shipperDetilsMasterCell.Split('~');
                                    string AddressLine1 = clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1;
                                    string City = clsF3.ClsOCRValues.ClsOCRValuesShipperCity;
                                    string State = clsF3.ClsOCRValues.ClsOCRValuesShipperState;
                                    string Zip = clsF3.ClsOCRValues.ClsOCRValuesShipperZip;
                                    string shipperDetilsMasterCellOCR = AddressLine1 + " " + City + " " + State + " " + Zip;
                                    string shipperDetilsMasterCellMst = arrshipperDetilsMasterCell[1] + " " + arrshipperDetilsMasterCell[2] + " " + arrshipperDetilsMasterCell[3] + " " + arrshipperDetilsMasterCell[4];
                                    string IsMatchAddress = MatchAddress(shipperDetilsMasterCellOCR, shipperDetilsMasterCellMst, "S");
                                    // Need to match name along with address added on 6-feb-2021
                                    string NameFinalDB = arrshipperDetilsMasterCell[0].ToString().ToUpper();
                                    NameFinalDB = Module1.func_RemoveSpecialWordsFromName(NameFinalDB);//NameFinalDB.Replace(" COMPANY", "").Replace(" CO", "").Replace(" INCORPORATED", "").Replace(" INC", "").Replace(" ENTERPRISES", "").Replace(" ENTERPRISE", "").Replace(" LLC", "");

                                    string NameFinalOCR = clsF3.ClsOCRValues.ClsOCRValuesShipperName.ToString().ToUpper();
                                    NameFinalOCR = Module1.func_RemoveSpecialWordsFromName(NameFinalOCR); //NameFinalOCR.Replace(" COMPANY", "").Replace(" CO", "").Replace(" INCORPORATED", "").Replace(" INC", "").Replace(" ENTERPRISES", "").Replace(" ENTERPRISE", "").Replace(" LLC", "");

                                    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(NameFinalDB.ToString()), Module1.func_RemoveSpecialCharacter(NameFinalOCR));

                                    if (objRet50.PercentageMatch >= 90)
                                    {
                                        IsMatchNameShipper = "Y";
                                    }
                                    if (IsMatchAddress == "Y") // && IsMatchNameShipper == "Y"
                                    {
                                        strLogic = "YES";
                                    }
                                    else
                                    {
                                        strLogic = "NO";
                                    }
                                }
                                else
                                {
                                    strLogic = "NO";
                                }
                            }
                            else
                            {
                                strLogic = "NA";
                            }
                        }
                        clsModule.OCR_Matched_MasterDB_KeyShipper = strLogic;
                        objStructF27.ContentsOfField = strLogic;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C12": //  Shipper Details OCR_matched_Search enagine
                        string LogicCell = clsModule.OCR_Matched_MasterDB_KeyShipper;
                        string strlogic1 = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "@@@@@")
                        {
                            if (!string.IsNullOrEmpty(LogicCell))
                            {
                                if (LogicCell == "YES")
                                {
                                    strlogic1 = "NA";
                                }
                                else
                                {
                                    // TO DO => Need to use OCR_matched_Search enagine
                                    if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperName) && clsF3.ClsOCRValues.ClsOCRValuesShipperName != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperName != "@@@@@")
                                    {
                                        DataTable dtSearchEngine = new DataTable();
                                        string strflag = "S";
                                        dtSearchEngine = GetDataFromSearchEngine(strflag);
                                        if (dtSearchEngine != null && dtSearchEngine.Rows.Count > 0)
                                        {
                                            strlogic1 = "YES";
                                            if (dtSearchEngine != null)
                                            {
                                                dtSearchEngine.Dispose();
                                                dtSearchEngine = null;
                                            }
                                        }
                                        else
                                        {
                                            strlogic1 = "NO";
                                        }
                                    }
                                    else
                                    {
                                        strlogic1 = "NO";
                                    }
                                }

                            }
                        }
                        else
                        {
                            strlogic1 = "NA";
                        }
                        //}
                        clsModule.OCR_matched_SearchenagineShipper = strlogic1;
                        objStructF27.ContentsOfField = strlogic1;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C007A": // Shipper_Code# Final
                        string ShipperCodeFinal = string.Empty;
                        string ShipperPickUpCode = string.Empty; // Added on 23-July-2021
                        string ShipperDetailsMasterDBKey = clsModule.OCR_Matched_MasterDB_KeyShipper;
                        string ShipperDetailsSearchEngine = clsModule.OCR_matched_SearchenagineShipper;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "@@@@@")
                        {
                            // Pickup => FRP001
                            if (!string.IsNullOrEmpty(ShipperDetailsMasterDBKey) && ShipperDetailsMasterDBKey == "YES")
                            {
                                if (clsModule.clsPickupValuesFRP001RAT == "Y" && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATCodeFinal))
                                {
                                    ShipperCodeFinal = clsModule.clsMasterValuesAPR001ShipperRATCodeFinal;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "5F";
                                }
                                else // Chcek for normal case Shipper and BS
                                {
                                    ShipperCodeFinal = clsModule.clsMasterValuesAPR001ShipperCode;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "5F";
                                }
                            }
                            // Master => ARP001
                            else if (!string.IsNullOrEmpty(ShipperDetailsSearchEngine) && ShipperDetailsSearchEngine == "YES")
                            {
                                if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "F") // Valid Code Found = GREEN 
                                {
                                    ShipperCodeFinal = ClsSearchEngin.ClsSearchEnginShipperCode;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "AD" || ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "AO")  // Alias Code Found = GREEN (AD = Alias DB , AO = Alias OBL)
                                {
                                    ShipperCodeFinal = ClsSearchEngin.ClsSearchEnginShipperCode;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "AF") // only address match then display name as OBL and code will be red only (ARP001 found one record and No alias data found)
                                {
                                    ShipperCodeFinal = ClsSearchEngin.ClsSearchEnginShipperCode;
                                    objStructF27.Flag = "1"; //"0";
                                    objStructF27.Remarks = "6F";
                                }
                                else // Multiple code found = RED OR Code Not found
                                {
                                    ShipperCodeFinal = ClsSearchEngin.ClsSearchEnginShipperCode;
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "6F";
                                }
                            }
                            // OCR Value
                            else
                            {
                                //if (clsModule.ShipperKMFlag == "T" && !string.IsNullOrEmpty(clsModule.ShipperKMFinalCode))
                                //{
                                //    ShipperCodeFinal = clsModule.ShipperKMFinalCode;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "7F";
                                //}
                                //else
                                //{
                                //    ShipperCodeFinal = clsModule.clsPickupValuesFRP001ShipperCode; //"*****";//clsModule.clsPickupValuesFRP001ShipperCode;
                                //    objStructF27.Flag = "0";
                                //    objStructF27.Remarks = "1F";
                                //}
                                ShipperCodeFinal = clsModule.clsPickupValuesFRP001ShipperCode; //"*****";//clsModule.clsPickupValuesFRP001ShipperCode;
                                objStructF27.Flag = "0";
                                objStructF27.Remarks = "1F";
                                ShipperPickUpCode = "T"; // Added on 23-July-2021

                            }
                        }
                        else
                        {
                            // Added on 22-July-2021 
                            ShipperCodeFinal = clsModule.clsPickupValuesFRP001ShipperCode;
                            objStructF27.Flag = "0";
                            objStructF27.Remarks = "1F";
                            clsModule.isShipperValidFlag = "F";
                            ShipperPickUpCode = "T"; // Added on 23-July-2021
                        }
                        #region "drop ShipperKM final details when shipper details not found from master/search enginee"
                        if (ShipperDetailsMasterDBKey != "YES" && ShipperDetailsSearchEngine != "YES" && ShipperPickUpCode == "T")
                        {
                            string ShipperDetails_PKP = clsModule.clsPickupValuesFRP001ShipperDetailsPickUp;
                            string spName = "USP_ShipperKM_GetData";
                            SqlParameter ShipperDetailsPKPParameter = new SqlParameter("@ShipperDetails_PKP", ShipperDetails_PKP);
                            SqlParameter[] parameters = new SqlParameter[] { ShipperDetailsPKPParameter };
                            DataTable dtShipperKM = GetShipperKM(spName, parameters);
                            if (dtShipperKM != null && dtShipperKM.Rows != null && dtShipperKM.Rows.Count == 1)
                            {
                                clsModule.ShipperKMFinalCode = dtShipperKM.Rows[0]["ShipCode"].ToString();
                                clsModule.ShipperKMFinalName = dtShipperKM.Rows[0]["ShipName_FNL"].ToString();
                                clsModule.ShipperKMFinalAddressLine = dtShipperKM.Rows[0]["ShipAddress_FNL"].ToString();
                                clsModule.ShipperKMFinalCity = dtShipperKM.Rows[0]["ShipCity_FNL"].ToString();
                                clsModule.ShipperKMFinalState = dtShipperKM.Rows[0]["ShipState_FNL"].ToString();
                                clsModule.ShipperKMFinalZip = dtShipperKM.Rows[0]["Ship_Zip_FNL"].ToString();
                                clsModule.ShipperKMFlag = "T";
                                clsModule.ShipperKMOBLName = dtShipperKM.Rows[0]["ShipName_OBL"].ToString();
                                clsModule.ShipperKMOBLAddressLine = dtShipperKM.Rows[0]["ShipAddress_OBL"].ToString();
                            }
                            else
                            {
                                clsModule.ShipperKMFlag = "F";
                            }
                        }
                        else
                        {
                            clsModule.ShipperKMFlag = "F";
                        }
                        if (clsModule.ShipperKMFlag == "T" && !string.IsNullOrEmpty(clsModule.ShipperKMFinalCode))
                        {
                            ShipperCodeFinal = clsModule.ShipperKMFinalCode;
                            objStructF27.Flag = "1";
                            objStructF27.Remarks = "7F";
                        }

                        #endregion
                        if (!string.IsNullOrEmpty(ShipperCodeFinal) && ShipperCodeFinal != "*****" && ShipperCodeFinal != "@@@@@")
                        {
                            ShipperCodeFinal = ShipperCodeFinal.PadLeft(ShipperCodeFinal.Length, '0');
                            objStructF27.ContentsOfField = ShipperCodeFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                            objStructF27.Flag = "0"; // added on 10-March-2021
                        }
                        clsModule.ShipperCodeFinal = ShipperCodeFinal;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET4_ShpCode"))
                        {
                            iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET4_ShpCode");
                        }
                        iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET4_ShpCode", ShipperCodeFinal);
                        // check code is valid or not
                        if (objStructF27.Flag == "0" && clsModule.isShipperValidFlag == "T")
                        {
                            clsModule.isShipperValidFlag = "F";
                        }

                        break;
                    case "C007": // Shipper Name Final
                        string ShipperNameFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "@@@@@")
                        {
                            // check OCR character 
                            string ShipperDetailsMasterDBKeyName = clsModule.OCR_Matched_MasterDB_KeyShipper;
                            string ShipperDetailsSearchEngineName = clsModule.OCR_matched_SearchenagineShipper;
                            string ShipperDetailsMaster = clsModule.clsMasterValuesAPR001ShipperDetails;
                            if (!string.IsNullOrEmpty(ShipperDetailsMasterDBKeyName) && ShipperDetailsMasterDBKeyName == "YES")
                            {
                                // STEP 1 => Check for RAT
                                if (clsModule.clsPickupValuesFRP001RAT == "Y")
                                {
                                    // if valid shipper code found in the master against shipper RAT bill then compare name with 90% match
                                    if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATDetails))
                                    {
                                        string[] arrshipperDetilsMasterRAT = clsModule.clsMasterValuesAPR001ShipperRATDetails.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterRAT[0]))
                                        {
                                            objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(arrshipperDetilsMasterRAT[0].ToString()), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperName));
                                            if (objRet50.PercentageMatch >= 90)
                                            {
                                                ShipperNameFinal = arrshipperDetilsMasterRAT[0].ToString();
                                                objStructF27.Flag = "1";
                                                objStructF27.Remarks = "5F";
                                            }
                                            else
                                            {
                                                // check the confidance level if it is greater than 90 then make it green
                                                //if (clsF3.ClsOCRValues.ShipperNameOCRWithHighConfidence == "Y")
                                                //{
                                                //    ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                                //    objStructF27.Flag = "1";
                                                //    objStructF27.Remarks = "1F";
                                                //}
                                                //else
                                                //{
                                                ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                                objStructF27.Flag = "0";
                                                objStructF27.Remarks = "1F";
                                                //}
                                            }
                                            //DBShipperNameFinal = "Y";
                                        }

                                    } // if NOT found for RAT
                                    else
                                    {
                                        //if (clsF3.ClsOCRValues.ShipperNameOCRWithHighConfidence == "Y")
                                        //{
                                        //    ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                        //    objStructF27.Flag = "1";
                                        //    objStructF27.Remarks = "1F";
                                        //}
                                        //else
                                        //{
                                        ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                        //}
                                    }
                                }
                                // STEP 2 => Check for BS => Blind Shipper , if we found blind shipper then check confidence string and make it green
                                else if (clsModule.clsPickupValuesFRP001BSMatched == "Y" || clsModule.clsPickupValuesFRP001BS == "Y")
                                {
                                    if (clsF3.ClsOCRValues.ClsOCRValuesShipperName != "*****" && !string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperName))
                                    {
                                        if (clsF3.ClsOCRValues.ShipperNameOCRWithHighConfidence == "Y")
                                        {
                                            ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName + "   " + "(BS)";
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "1F";
                                        }
                                        else
                                        {
                                            ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName + "   " + "(BS)";
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                    else
                                    {
                                        //if (clsF3.ClsOCRValues.ShipperNameOCRWithHighConfidence == "Y")
                                        //{
                                        //    ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                        //    objStructF27.Flag = "1";
                                        //    objStructF27.Remarks = "1F";
                                        //}
                                        //else
                                        //{
                                        ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                        //}
                                    }

                                }
                                // STEP 3 => Check for normal case
                                else
                                {
                                    string[] arrshipperDetilsMasterCell = ShipperDetailsMaster.Split('~');
                                    string NameFinalDB = arrshipperDetilsMasterCell[0].ToString().ToUpper();
                                    string NameFinalDBOriginal = arrshipperDetilsMasterCell[0].ToString().ToUpper();
                                    NameFinalDB = Module1.func_RemoveSpecialWordsFromName(NameFinalDB);//NameFinalDB.Replace(" COMPANY", "").Replace(" CO", "").Replace(" INCORPORATED", "").Replace(" INC", "").Replace(" ENTERPRISES", "").Replace(" ENTERPRISE", "").Replace(" LLC", "");

                                    string NameFinalOCR = clsF3.ClsOCRValues.ClsOCRValuesShipperName.ToString().ToUpper();
                                    string NameFinalOCROriginal = clsF3.ClsOCRValues.ClsOCRValuesShipperName.ToString().ToUpper();
                                    NameFinalOCR = Module1.func_RemoveSpecialWordsFromName(NameFinalOCR); //NameFinalOCR.Replace(" COMPANY", "").Replace(" CO", "").Replace(" INCORPORATED", "").Replace(" INC", "").Replace(" ENTERPRISES", "").Replace(" ENTERPRISE", "").Replace(" LLC", "");


                                    if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[0]))
                                    {
                                        objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(NameFinalDB.ToString()), Module1.func_RemoveSpecialCharacter(NameFinalOCR));
                                        if (objRet50.PercentageMatch >= 90)
                                        {
                                            ShipperNameFinal = arrshipperDetilsMasterCell[0].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                        else
                                        {
                                            // check alias for name in master table if single record found then drop master name
                                            //string AddressLine1 = Regex.Split(clsModule.clsMasterValuesAPR001ShipperDetailsWithPipe, @"\|\|")[0];
                                            //DataTable dtAlias = GetAliasDataFromMaster(As400_ConnectionString, arrshipperDetilsMasterCell[3].ToString(), arrshipperDetilsMasterCell[4].ToString(), AddressLine1.ToString(), arrshipperDetilsMasterCell[2].ToString());
                                            string CustomerCode = string.Empty;
                                            if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperCode) && clsModule.clsMasterValuesAPR001ShipperCode != "*****")
                                            {
                                                CustomerCode = "'" + clsModule.clsMasterValuesAPR001ShipperCode.ToString().Trim() + "'";

                                                DataTable dtAlias = clsSearchCustomer.GetMatchedAliasFromARP033(CustomerCode, clsF3.ClsOCRValues.ClsOCRValuesShipperName);
                                                if (dtAlias != null && dtAlias.Rows.Count == 1)
                                                {
                                                    ShipperNameFinal = dtAlias.Rows[0]["NANAME"].ToString();
                                                    objStructF27.Flag = "1";
                                                    objStructF27.Remarks = "5F";
                                                }
                                                // check alias for name in master table if multiple record found then drop OBL name
                                                else
                                                {
                                                    string DBName = arrshipperDetilsMasterCell[0].Trim().ToUpper();
                                                    string OCRName = clsF3.ClsOCRValues.ClsOCRValuesShipperName.Trim().ToUpper();
                                                    //string isWordFound = string.Empty;
                                                    int isWordFound = WordMatchFromName(DBName, OCRName);
                                                    if (isWordFound == 1)
                                                    {
                                                        ShipperNameFinal = arrshipperDetilsMasterCell[0];
                                                        objStructF27.Flag = "1";
                                                        objStructF27.Remarks = "5F";
                                                        // Above code commented and added below on 30-July-2021 to block one word logic
                                                        //objStructF27.Flag = "0";
                                                        //objStructF27.Remarks = "WF";
                                                    }
                                                    else if (isWordFound > 1)
                                                    {
                                                        ShipperNameFinal = arrshipperDetilsMasterCell[0];
                                                        objStructF27.Flag = "1";
                                                        objStructF27.Remarks = "5F";
                                                    }
                                                    else
                                                    {
                                                        // 1) Take orignal Names of DB and OCR (with out removing words) 2) Remove space 3) percentage match >=90 4) Shipper Name (Master DB) => Green with flag = 1 and remark = 5F ELSE drop OBL Name flag = 0 and remark = 1F 
                                                        objRet50 = ObjSbr.IQSBR050(NameFinalDBOriginal.ToString().Trim().Replace(" ", ""), NameFinalOCROriginal.ToString().Trim().Replace(" ", ""));
                                                        if (objRet50.PercentageMatch >= 90)
                                                        {
                                                            ShipperNameFinal = arrshipperDetilsMasterCell[0].ToString();
                                                            objStructF27.Flag = "1";
                                                            objStructF27.Remarks = "5F";
                                                        }
                                                        else
                                                        {
                                                            ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                                            objStructF27.Flag = "0";
                                                            objStructF27.Remarks = "1F";
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                                objStructF27.Flag = "0";
                                                objStructF27.Remarks = "1F";
                                            }

                                        }
                                    }
                                }
                            }
                            else if (!string.IsNullOrEmpty(ShipperDetailsSearchEngineName) && ShipperDetailsSearchEngineName == "YES")

                            {
                                #region Commented Code
                                //if (ClsSearchEngin.ClsSearchEnginAliseShipperName != "Y")
                                //{
                                //    ShipperNameFinal = ClsSearchEngin.ClsSearchEnginShipperName;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                //else
                                //{
                                //    // Compare search engine name with OBL if it matched then dropped serach engine name else use OCR name and make both green
                                //    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(ClsSearchEngin.ClsSearchEnginShipperName.ToString()), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperName));
                                //    if (objRet50.PercentageMatch >= 90)
                                //    {
                                //        ShipperNameFinal = ClsSearchEngin.ClsSearchEnginShipperName;
                                //        objStructF27.Flag = "1";
                                //        objStructF27.Remarks = "6F";

                                //    }
                                //    else
                                //    {
                                //        ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                //        objStructF27.Flag = "1";
                                //        objStructF27.Remarks = "6F";
                                //    }
                                //}
                                #endregion

                                if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "F") // Valid Code Found = GREEN 
                                {
                                    ShipperNameFinal = ClsSearchEngin.ClsSearchEnginShipperName;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "AD")
                                {
                                    ShipperNameFinal = ClsSearchEngin.ClsSearchEnginAliseShipperNameARP033;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "AO")  // Alias Code Found = GREEN
                                {
                                    // Check DB Name and OBL Name if any one word match then make it green
                                    string DBName = ClsSearchEngin.ClsSearchEnginShipperName.Trim().ToUpper();
                                    string OCRName = clsF3.ClsOCRValues.ClsOCRValuesShipperName.Trim().ToUpper();
                                    //string isWordFound = string.Empty;
                                    int isWordFound = WordMatchFromName(DBName, OCRName);
                                    if (isWordFound == 1)
                                    {
                                        ShipperNameFinal = ClsSearchEngin.ClsSearchEnginShipperName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                        // Above code commented and added below on 30-July-2021 to block one word logic
                                        //objStructF27.Flag = "0";
                                        //objStructF27.Remarks = "WF";
                                    }
                                    else if (isWordFound > 1)
                                    {
                                        ShipperNameFinal = ClsSearchEngin.ClsSearchEnginShipperName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                    }
                                    else
                                    {
                                        // 1) Take orignal Names of DB and OCR (with out removing words) 2) Remove space 3) percentage match >=90 4) Shipper Name (Search engine) => Green with flag = 1 and remark = 6F ELSE drop OBL Name flag = 0 and remark = 1F 
                                        objRet50 = ObjSbr.IQSBR050(DBName.ToString().Trim().Replace(" ", ""), OCRName.ToString().Trim().Replace(" ", ""));
                                        if (objRet50.PercentageMatch >= 90)
                                        {
                                            ShipperNameFinal = ClsSearchEngin.ClsSearchEnginShipperName;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "6F";
                                        }
                                        else
                                        {
                                            ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                }
                                else if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "AF") // only address match then display name as OBL and code will be red only (ARP001 found one record and No alias data found)
                                {
                                    string DBName = ClsSearchEngin.ClsSearchEnginShipperName.ToUpper();
                                    string OCRName = clsF3.ClsOCRValues.ClsOCRValuesShipperName.ToUpper();
                                    //string isWordFound = string.Empty;
                                    int isWordFound = WordMatchFromName(DBName, OCRName);
                                    if (isWordFound == 1)
                                    {
                                        ShipperNameFinal = ClsSearchEngin.ClsSearchEnginShipperName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                        // Above code commented and added below on 30-July-2021 to block one word logic
                                        //objStructF27.Flag = "0";
                                        //objStructF27.Remarks = "WF";
                                    }
                                    else if (isWordFound > 1)
                                    {
                                        ShipperNameFinal = ClsSearchEngin.ClsSearchEnginShipperName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                    }
                                    else
                                    {
                                        // 1) Take orignal Names of DB and OCR (with out removing words) 2) Remove space 3) percentage match >=90 4) Shipper Name (Search engine) => Green with flag = 1 and remark = 6F ELSE drop OBL Name flag = 0 and remark = 1F 
                                        objRet50 = ObjSbr.IQSBR050(DBName.ToString().Trim().Replace(" ", ""), OCRName.ToString().Trim().Replace(" ", ""));
                                        if (objRet50.PercentageMatch >= 90)
                                        {
                                            ShipperNameFinal = ClsSearchEngin.ClsSearchEnginShipperName;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "6F";
                                        }
                                        else
                                        {
                                            ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                    //}
                                }
                                else // Multiple code found = RED
                                {
                                    ShipperNameFinal = ClsSearchEngin.ClsSearchEnginShipperName;
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "6F";
                                }
                            }
                            else
                            {
                                // If data is not found from master and serch enginee then match data with pick up and make it green
                                objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.clsPickupValuesFRP001ShipperName), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperName));
                                if (objRet50.PercentageMatch >= 90 && clsModule.ShipperKMFlag == "F")
                                {
                                    ShipperNameFinal = clsModule.clsPickupValuesFRP001ShipperName;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "1F";
                                }
                                else
                                {
                                    //ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                    //objStructF27.Flag = "0";
                                    //objStructF27.Remarks = "1F";
                                    if (clsModule.ShipperKMFlag == "T" && !string.IsNullOrEmpty(clsModule.ShipperKMFinalName)) // Added on 13-July-2021
                                    {
                                        // ShipperKM OBL Name should be 50% match with OCR name
                                        objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.ShipperKMOBLName), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperName));
                                        if (objRet50.PercentageMatch > 50)
                                        {
                                            ShipperNameFinal = clsModule.ShipperKMFinalName;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "7F";
                                        }
                                        else
                                        {
                                            ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                    else
                                    {
                                        ShipperNameFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperName;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }
                            }
                        }
                        else
                        {
                            clsModule.isShipperValidFlag = "F";

                        }
                        if (!string.IsNullOrEmpty(ShipperNameFinal))
                        {
                            objStructF27.ContentsOfField = ShipperNameFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                            objStructF27.Flag = "0"; // added on 10-March-2021
                        }
                        clsModule.clsFinalValuesShipperName = ShipperNameFinal;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        // check valid code
                        if (objStructF27.Flag == "0" && clsModule.isShipperValidFlag == "T")
                        {
                            clsModule.isShipperValidFlag = "F";
                        }
                        break;

                    case "C008": // Shipper address Line Final
                        string ShipperAddressLineFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "@@@@@")
                        {
                            string ShipperDetailsMasterDBKeyAddressLine = clsModule.OCR_Matched_MasterDB_KeyShipper; ;
                            string ShipperDetailsSearchEngineAddressLine = clsModule.OCR_matched_SearchenagineShipper;
                            string ShipperDetailsMasterAddress = clsModule.clsMasterValuesAPR001ShipperDetails;
                            if (!string.IsNullOrEmpty(ShipperDetailsMasterDBKeyAddressLine) && ShipperDetailsMasterDBKeyAddressLine == "YES")
                            {
                                // STEP 1 - Check for RAT
                                if (clsModule.clsPickupValuesFRP001RAT == "Y")
                                {
                                    // if valid shipper code found in the master against shipper RAT bill
                                    if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATDetails))
                                    {
                                        string[] arrshipperDetilsMasterRAT = clsModule.clsMasterValuesAPR001ShipperRATDetails.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterRAT[1]))
                                        {
                                            ShipperAddressLineFinal = clsModule.clsMasterValuesAPR001ShipperDetailsWithPipe; //arrshipperDetilsMasterRAT[1].ToString(); 
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }

                                    }
                                    else
                                    {
                                        ShipperAddressLineFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }
                                // STEP 2 - Check for Normal case and BS
                                else
                                {
                                    string[] arrshipperDetilsMasterCell = ShipperDetailsMasterAddress.Trim().Split('~');
                                    if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[1]))
                                    {
                                        ShipperAddressLineFinal = clsModule.clsMasterValuesAPR001ShipperDetailsWithPipe; //arrshipperDetilsMasterCell[1].ToString(); 
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                            }
                            else if (!string.IsNullOrEmpty(ShipperDetailsSearchEngineAddressLine) && ShipperDetailsSearchEngineAddressLine == "YES")

                            {

                                ShipperAddressLineFinal = ClsSearchEngin.ClsSearchEnginShipperDetailsWithPipe;
                                objStructF27.Flag = "1";
                                objStructF27.Remarks = "6F";

                                #region Commented Code
                                //if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "F") // Valid Code Found = GREEN 
                                //{
                                //    ShipperAddressLineFinal = ClsSearchEngin.ClsSearchEnginShipperDetailsWithPipe; 
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                //else if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "A")  // Alias Code Found = GREEN
                                //{
                                //    ShipperAddressLineFinal = ClsSearchEngin.ClsSearchEnginShipperDetailsWithPipe;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                //else // Multiple code found = RED
                                //{
                                //    ShipperAddressLineFinal = ClsSearchEngin.ClsSearchEnginShipperDetailsWithPipe;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                #endregion
                            }
                            else
                            {
                                // If data is not found from master and serch enginee then match data with pick up and make it green
                                objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.clsPickupValuesFRP001ShipperAddress1), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1));
                                if (objRet50.PercentageMatch >= 70 && clsModule.ShipperKMFlag == "F")
                                {
                                    ShipperAddressLineFinal = clsModule.clsPickupValuesFRP001ShipperAddress1;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "1F";
                                }
                                else
                                {
                                    if (clsModule.ShipperKMFlag == "T" && !string.IsNullOrEmpty(clsModule.ShipperKMFinalAddressLine)) // Added on 13-July-2021
                                    {
                                        // ShipperKM OBL Name should be 50% match with OCR name
                                        objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.ShipperKMOBLAddressLine), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1));
                                        if (objRet50.PercentageMatch > 50)
                                        {
                                            ShipperAddressLineFinal = clsModule.ShipperKMFinalAddressLine;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "7F";
                                        }
                                        else
                                        {
                                            ShipperAddressLineFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }

                                    }
                                    else
                                    {
                                        ShipperAddressLineFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }
                            }
                        }
                        else
                        {
                            clsModule.isShipperValidFlag = "F";
                        }
                        if (!string.IsNullOrEmpty(ShipperAddressLineFinal))
                        {
                            objStructF27.ContentsOfField = ShipperAddressLineFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                            objStructF27.Flag = "0"; // added on 10-March-2021
                        }
                        clsModule.clsFinalValuesShipperAddress1 = ShipperAddressLineFinal;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        // check valid code
                        if (objStructF27.Flag == "0" && clsModule.isShipperValidFlag == "T")
                        {
                            clsModule.isShipperValidFlag = "F";
                        }
                        break;

                    case "C08A": // City Final
                        string ShipperCityFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "@@@@@")
                        {
                            string ShipperDetailsMasterDBKeyCity = clsModule.OCR_Matched_MasterDB_KeyShipper;
                            string ShipperDetailsSearchEngineCity = clsModule.OCR_matched_SearchenagineShipper;
                            string ShipperDetailsMasterCity = clsModule.clsMasterValuesAPR001ShipperDetails;
                            if (!string.IsNullOrEmpty(ShipperDetailsMasterDBKeyCity) && ShipperDetailsMasterDBKeyCity == "YES")
                            {
                                // STEP 1 - check shipper => RAT
                                if (clsModule.clsPickupValuesFRP001RAT == "Y")
                                {
                                    // if valid shipper code found in the master against shipper RAT bill
                                    if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATDetails))
                                    {
                                        string[] arrshipperDetilsMasterRAT = clsModule.clsMasterValuesAPR001ShipperRATDetails.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterRAT[2]))
                                        {
                                            ShipperCityFinal = arrshipperDetilsMasterRAT[2].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }

                                    }
                                    else
                                    {
                                        ShipperCityFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperCity;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }
                                // STEP 2 - Check for Normal case and BS
                                else
                                {
                                    string[] arrshipperDetilsMasterCell = ShipperDetailsMasterCity.Split('~');
                                    if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[2]))
                                    {
                                        ShipperCityFinal = arrshipperDetilsMasterCell[2].ToString();
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                            }
                            else if (!string.IsNullOrEmpty(ShipperDetailsSearchEngineCity) && ShipperDetailsSearchEngineCity == "YES")
                            {
                                ShipperCityFinal = ClsSearchEngin.ClsSearchEnginShipperCity;
                                objStructF27.Flag = "1";
                                objStructF27.Remarks = "6F";

                                #region Commented code
                                //if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "F") // Valid Code Found = GREEN 
                                //{
                                //    ShipperCityFinal = ClsSearchEngin.ClsSearchEnginShipperCity;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                //else if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "A")  // Alias Code Found = GREEN
                                //{
                                //    ShipperCityFinal = ClsSearchEngin.ClsSearchEnginShipperCity;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                //else // Multiple code found = RED
                                //{
                                //    ShipperCityFinal = ClsSearchEngin.ClsSearchEnginShipperCity;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                #endregion
                            }
                            else
                            {
                                // If data is not found from master and serch enginee then match data with pick up and make it green
                                objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.clsPickupValuesFRP001ShipperCity), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperCity));
                                if (objRet50.PercentageMatch == 100 && clsModule.ShipperKMFlag == "F")
                                {
                                    ShipperCityFinal = clsModule.clsPickupValuesFRP001ShipperCity;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "1F";
                                }
                                else
                                {
                                    if (clsModule.ShipperKMFlag == "T" && !string.IsNullOrEmpty(clsModule.ShipperKMFinalCity)) // Added on 13-July-2021
                                    {
                                        ShipperCityFinal = clsModule.ShipperKMFinalCity;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "7F";
                                    }
                                    else
                                    {
                                        ShipperCityFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperCity;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }

                            }
                        }
                        else
                        {
                            clsModule.isShipperValidFlag = "F";

                        }
                        if (!string.IsNullOrEmpty(ShipperCityFinal))
                        {
                            objStructF27.ContentsOfField = ShipperCityFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                            objStructF27.Flag = "0"; // added on 10-March-2021
                        }
                        clsModule.clsFinalValuesShipperCity = ShipperCityFinal;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        // check valid code
                        if (objStructF27.Flag == "0" && clsModule.isShipperValidFlag == "T")
                        {
                            clsModule.isShipperValidFlag = "F";
                        }
                        break;

                    case "C08B": // State Final
                        string ShipperStateFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "@@@@@")
                        {
                            string ShipperDetailsMasterDBKeyState = clsModule.OCR_Matched_MasterDB_KeyShipper;
                            string ShipperDetailsSearchEngineSate = clsModule.OCR_matched_SearchenagineShipper;
                            string ShipperDetailsMasterState = clsModule.clsMasterValuesAPR001ShipperDetails;
                            if (!string.IsNullOrEmpty(ShipperDetailsMasterDBKeyState) && ShipperDetailsMasterDBKeyState == "YES")
                            {
                                // STEP 1 -  check shipper => RAT
                                if (clsModule.clsPickupValuesFRP001RAT == "Y")
                                {
                                    // if valid shipper code found in the master against shipper RAT bill
                                    if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATDetails))
                                    {
                                        string[] arrshipperDetilsMasterRAT = clsModule.clsMasterValuesAPR001ShipperRATDetails.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterRAT[3]))
                                        {
                                            ShipperStateFinal = arrshipperDetilsMasterRAT[3].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }

                                    }
                                    else
                                    {
                                        ShipperStateFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperState;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }
                                else
                                {
                                    string[] arrshipperDetilsMasterCell = ShipperDetailsMasterState.Split('~');
                                    if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[3]))
                                    {
                                        ShipperStateFinal = arrshipperDetilsMasterCell[3].ToString();
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                            }
                            else if (!string.IsNullOrEmpty(ShipperDetailsSearchEngineSate) && ShipperDetailsSearchEngineSate == "YES")
                            {
                                ShipperStateFinal = ClsSearchEngin.ClsSearchEnginShipperState;
                                objStructF27.Flag = "1";
                                objStructF27.Remarks = "6F";

                                #region Commented code
                                //if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "F") // Valid Code Found = GREEN 
                                //{
                                //    ShipperStateFinal = ClsSearchEngin.ClsSearchEnginShipperState;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                //else if (ClsSearchEngin.ClsSearchCommonStatusFlagShipper == "A")  // Alias Code Found = GREEN
                                //{
                                //    ShipperStateFinal = ClsSearchEngin.ClsSearchEnginShipperState;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                //else // Multiple code found = RED
                                //{
                                //    ShipperStateFinal = ClsSearchEngin.ClsSearchEnginShipperState;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                #endregion
                            }
                            else
                            {
                                // If data is not found from master and serch enginee then match data with pick up and make it green
                                objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.clsPickupValuesFRP001ShipperState), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperState));
                                if (objRet50.PercentageMatch == 100 && clsModule.ShipperKMFlag == "F")
                                {
                                    ShipperStateFinal = clsModule.clsPickupValuesFRP001ShipperState;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "1F";
                                }
                                else
                                {
                                    if (clsModule.ShipperKMFlag == "T" && !string.IsNullOrEmpty(clsModule.ShipperKMFinalState)) // Added on 13-July-2021
                                    {
                                        ShipperStateFinal = clsModule.ShipperKMFinalState;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "7F";
                                    }
                                    else
                                    {
                                        ShipperStateFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperState;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }

                            }
                        }
                        else
                        {
                            clsModule.isShipperValidFlag = "F";

                        }
                        if (!string.IsNullOrEmpty(ShipperStateFinal))
                        {
                            objStructF27.ContentsOfField = ShipperStateFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                            objStructF27.Flag = "0"; // added on 10-March-2021
                        }
                        clsModule.clsFinalValuesShipperState = ShipperStateFinal;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        // check valid code
                        if (objStructF27.Flag == "0" && clsModule.isShipperValidFlag == "T")
                        {
                            clsModule.isShipperValidFlag = "F";
                        }
                        break;

                    case "C08C": // Zip Final
                        string ShipperZipFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "@@@@@")
                        {
                            string ShipperDetailsMasterDBKeyZip = clsModule.OCR_Matched_MasterDB_KeyShipper;
                            string ShipperDetailsSearchEngineZip = clsModule.OCR_matched_SearchenagineShipper;
                            string ShipperDetailsMasterZip = clsModule.clsMasterValuesAPR001ShipperDetails;
                            if (!string.IsNullOrEmpty(ShipperDetailsMasterDBKeyZip) && ShipperDetailsMasterDBKeyZip == "YES")
                            {
                                if (clsModule.clsPickupValuesFRP001RAT == "Y")
                                {
                                    // if valid shipper code found in the master against shipper RAT bill
                                    if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATDetails))
                                    {
                                        string[] arrshipperDetilsMasterRAT = clsModule.clsMasterValuesAPR001ShipperRATDetails.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterRAT[4]))
                                        {
                                            ShipperZipFinal = arrshipperDetilsMasterRAT[4].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }

                                    }
                                    else
                                    {
                                        ShipperZipFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperZip;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }
                                else
                                {
                                    string[] arrshipperDetilsMasterCell = ShipperDetailsMasterZip.Split('~');
                                    if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[4]))
                                    {
                                        ShipperZipFinal = arrshipperDetilsMasterCell[4].ToString();
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                            }
                            else if (!string.IsNullOrEmpty(ShipperDetailsSearchEngineZip) && ShipperDetailsSearchEngineZip == "YES")
                            {
                                ShipperZipFinal = ClsSearchEngin.ClsSearchEnginShipperZip;
                                objStructF27.Flag = "1";
                                objStructF27.Remarks = "6F";
                            }
                            else
                            {
                                // If data is not found from master and serch enginee then match data with pick up and make it green
                                objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.clsPickupValuesFRP001ShipperZip), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperZip));
                                if (objRet50.PercentageMatch == 100 && clsModule.ShipperKMFlag == "F")
                                {
                                    ShipperZipFinal = clsModule.clsPickupValuesFRP001ShipperZip;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "1F";
                                }
                                else
                                {
                                    if (clsModule.ShipperKMFlag == "T" && !string.IsNullOrEmpty(clsModule.ShipperKMFinalZip)) // Added on 13-July-2021
                                    {
                                        ShipperZipFinal = clsModule.ShipperKMFinalZip;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "7F";
                                    }
                                    else
                                    {
                                        ShipperZipFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperZip;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }
                            }
                        }
                        else
                        {
                            clsModule.isShipperValidFlag = "F";
                        }
                        if (!string.IsNullOrEmpty(ShipperZipFinal))
                        {
                            objStructF27.ContentsOfField = ShipperZipFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                            objStructF27.Flag = "0"; // added on 10-March-2021
                        }
                        clsModule.clsFinalValuesShipperZip = ShipperZipFinal;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        // check valid code
                        if (objStructF27.Flag == "0" && clsModule.isShipperValidFlag == "T")
                        {
                            clsModule.isShipperValidFlag = "F";
                        }
                        break;

                    case "C08D": // Shipper Telephone Number Final
                        string ShipperTelFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim() != "@@@@@")
                        {
                            string ShipperDetailsMasterDBKeyTelePhoneNo = clsModule.OCR_Matched_MasterDB_KeyShipper;
                            string ShipperDetailsSearchEngineTelePhoneNo = clsModule.OCR_matched_SearchenagineShipper;
                            // Master Engine
                            if (!string.IsNullOrEmpty(ShipperDetailsMasterDBKeyTelePhoneNo) && ShipperDetailsMasterDBKeyTelePhoneNo == "YES")
                            {
                                if (clsModule.clsPickupValuesFRP001RAT == "Y")
                                {
                                    // if valid shipper code found in the master against shipper RAT bill
                                    if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperRATDetails))
                                    {

                                        objRet58 = ObjSbr.IQSBR058(Module1.func_RemoveSpecialCharacter(clsModule.clsMasterValuesAPR001ShipperRATTelePhoneNo), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo), true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
                                        if (objRet58.CharChanged <= 1)
                                        {
                                            ShipperTelFinal = clsModule.clsMasterValuesAPR001ShipperRATTelePhoneNo;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                        //OCR
                                        else
                                        {
                                            if (clsF3.ClsOCRValues.ShipperTelephoneOCRWithHighConfidence == "Y" && clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo != "*****")
                                            {
                                                ShipperTelFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                                                objStructF27.Flag = "1";
                                                objStructF27.Remarks = "1F";
                                            }
                                            else
                                            {
                                                ShipperTelFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                                                objStructF27.Flag = "0";
                                                objStructF27.Remarks = "1F";
                                            }
                                        }

                                    }
                                    else
                                    {
                                        if (clsF3.ClsOCRValues.ShipperTelephoneOCRWithHighConfidence == "Y" && clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo != "*****")
                                        {
                                            ShipperTelFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "1F";
                                        }
                                        else
                                        {
                                            ShipperTelFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                }
                                else
                                {
                                    // check and comapre on OBL
                                    objRet58 = ObjSbr.IQSBR058(Module1.func_RemoveSpecialCharacter(clsModule.clsMasterValuesAPR001ShipperTelePhoneNo), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo), true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
                                    if (objRet58.CharChanged <= 1)
                                    {
                                        ShipperTelFinal = clsModule.clsMasterValuesAPR001ShipperTelePhoneNo;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                    // OCR
                                    else
                                    {
                                        if (clsF3.ClsOCRValues.ShipperTelephoneOCRWithHighConfidence == "Y" && clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo != "*****")
                                        {
                                            ShipperTelFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "1F";
                                        }
                                        else
                                        {
                                            ShipperTelFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                }
                            }
                            // Search Engine 
                            else if (!string.IsNullOrEmpty(ShipperDetailsSearchEngineTelePhoneNo) && ShipperDetailsSearchEngineTelePhoneNo == "YES")
                            {
                                objRet58 = ObjSbr.IQSBR058(Module1.func_RemoveSpecialCharacter(ClsSearchEngin.ClsSearchEnginShipperTelePhoneNo), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo), true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
                                if (objRet58.CharChanged <= 1)
                                {
                                    ShipperTelFinal = ClsSearchEngin.ClsSearchEnginShipperTelePhoneNo;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                // OCR
                                else
                                {
                                    if (clsF3.ClsOCRValues.ShipperTelephoneOCRWithHighConfidence == "Y" && clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo != "*****")
                                    {
                                        ShipperTelFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "1F";
                                    }
                                    else
                                    {
                                        ShipperTelFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }
                            }
                            // OCR
                            else
                            {
                                if (clsF3.ClsOCRValues.ShipperTelephoneOCRWithHighConfidence == "Y")
                                {
                                    ShipperTelFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "1F";
                                }
                                else
                                {
                                    ShipperTelFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "1F";
                                }
                            }
                        }
                        //}
                        if (!string.IsNullOrEmpty(ShipperTelFinal) && ShipperTelFinal != "*****" && ShipperTelFinal != "@@@@@" && ShipperTelFinal != "0")
                        {
                            objStructF27.ContentsOfField = ShipperTelFinal;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo))
                            {
                                ShipperTelFinal = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                                objStructF27.ContentsOfField = ShipperTelFinal;
                                objStructF27.Flag = "0";
                                objStructF27.Remarks = "1F";
                            }
                            else
                            {
                                objStructF27.ContentsOfField = "*****";
                            }
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        // Added on 24-March-2021 => Added valid shipper flag in PublicComponents for line items
                        if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET2_ShipperValidFlag"))
                        {
                            iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET2_ShipperValidFlag");
                        }
                        iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET2_ShipperValidFlag", clsModule.isShipperValidFlag);
                        break;

                    case "C011": //  Consignee Code# => Pickup FRP001
                        string ConsigneeCodePickup = clsModule.clsPickupValuesFRP001ConsigneeCode;
                        if (!string.IsNullOrEmpty(ConsigneeCodePickup))
                        {
                            ConsigneeCodePickup = ConsigneeCodePickup.PadLeft(7, '0');
                            objStructF27.ContentsOfField = ConsigneeCodePickup;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }

                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C022": // Consignee Details (~Sep) => Pickup => Name~Add~City~State~Zip~TelNo
                        string ConsigneeDetailsPickup = string.Empty;
                        if (!string.IsNullOrEmpty(clsModule.clsPickupValuesFRP001ConsigneeCode))
                        {
                            ConsigneeDetailsPickup = clsModule.clsPickupValuesFRP001ConsigneeName + "~" + clsModule.clsPickupValuesFRP001ConsigneeAddress1 + "~" + clsModule.clsPickupValuesFRP001ConsigneeCity + "~" + clsModule.clsPickupValuesFRP001ConsigneeState + "~" + clsModule.clsPickupValuesFRP001ConsigneeZip;
                        }
                        if (!string.IsNullOrEmpty(ConsigneeDetailsPickup))
                        {
                            objStructF27.ContentsOfField = ConsigneeDetailsPickup; ;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C099": // Consignee Master KeyField =>  Code_Pickup/TelNo_OCR
                        
                          if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "@@@@@")
                            {
                                string strConsigneeCodePickup = clsModule.clsPickupValuesFRP001ConsigneeCode;
                                string strRAT = clsModule.clsPickupValuesFRP001HAT;
                                if (!string.IsNullOrEmpty(strConsigneeCodePickup))
                                {
                                    string strConsigneeCodeMasterOrTelNoOCR = string.Empty;
                                    string strConsigneeDetailsMaster = string.Empty;
                                    string strflag = "C";
                                    GetShipperDetailsMaster(strConsigneeCodePickup, strflag, strRAT, ref strConsigneeCodeMasterOrTelNoOCR, ref strConsigneeDetailsMaster);
                                    string ShipperCodeMasterOrTelOCR = strConsigneeCodeMasterOrTelNoOCR;
                                    if (!string.IsNullOrEmpty(ShipperCodeMasterOrTelOCR))
                                    {
                                        objStructF27.ContentsOfField = ShipperCodeMasterOrTelOCR;
                                    }
                                    else
                                    {
                                        objStructF27.ContentsOfField = "@@@@@";
                                    }
                                }
                            }
                            else
                            {
                                objStructF27.ContentsOfField = "NA";
                            }

                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C100": // Consignee Details (~Sep) =>  Master
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "@@@@@")
                            {
                                string strRAT = clsModule.clsPickupValuesFRP001HAT;
                                string strConsigneeCodePickupDet = clsModule.clsPickupValuesFRP001ConsigneeCode;
                                if (!string.IsNullOrEmpty(strConsigneeCodePickupDet))
                                {
                                    string strConsigneeCodeMasterOrTelNoOCR = string.Empty;
                                    string strConsigneeDetailsMaster = string.Empty;
                                    string strflag = "C";
                                    GetShipperDetailsMaster(strConsigneeCodePickupDet, strflag, strRAT, ref strConsigneeCodeMasterOrTelNoOCR, ref strConsigneeDetailsMaster);
                                    string ConsigneeDetailMaster = strConsigneeDetailsMaster;
                                    if (!string.IsNullOrEmpty(ConsigneeDetailMaster))
                                    {
                                        objStructF27.ContentsOfField = ConsigneeDetailMaster;
                                    }
                                    else
                                    {
                                        objStructF27.ContentsOfField = "@@@@@";
                                    }
                                }
                            }
                            else
                            {
                                objStructF27.ContentsOfField = "NA";
                            }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C111": // Consignee Details OCR_Matched_MasterDB_Key => Yes/No
                        string strLogicConsignee = string.Empty;
                        string IsMatchNameConsignee = string.Empty;
                            // check for HAT
                            if ((clsModule.clsPickupValuesFRP001HAT == "Y"))
                            {
                                if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATDetails))
                                {
                                    strLogicConsignee = "YES";
                                }
                                else
                                {
                                    strLogicConsignee = "NO";
                                }
                            }
                            // chcek for normal case
                            else
                            {
                                if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "@@@@@")
                                {
                                    string ConsigneeDetilsMasterCell = clsModule.clsMasterValuesAPR001ConsigneeDetails;
                                    if (!string.IsNullOrEmpty(ConsigneeDetilsMasterCell) && ConsigneeDetilsMasterCell != "@@@@@")
                                    {
                                        string[] arrMasterDetailsCell = ConsigneeDetilsMasterCell.Split('~');
                                        string AddressLine1 = clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1;
                                        string City = clsF3.ClsOCRValues.ClsOCRValuesConsigneeCity;
                                        string State = clsF3.ClsOCRValues.ClsOCRValuesConsigneeState;
                                        string Zip = clsF3.ClsOCRValues.ClsOCRValuesConsigneeZip;
                                        string ConsigneeDetilsMasterCellOCR = AddressLine1 + " " + City + " " + State + " " + Zip;
                                        string ConsigneeDetilsMasterCellMst = arrMasterDetailsCell[1] + " " + arrMasterDetailsCell[2] + " " + arrMasterDetailsCell[3] + " " + arrMasterDetailsCell[4];
                                        string IsMatchAddress = MatchAddress(ConsigneeDetilsMasterCellOCR, ConsigneeDetilsMasterCellMst, "S");
                                        // Need to match name along with address added on 6-feb-2021
                                        string NameFinalDB = arrMasterDetailsCell[0].ToString().ToUpper();
                                        NameFinalDB = Module1.func_RemoveSpecialWordsFromName(NameFinalDB);//NameFinalDB.Replace(" COMPANY", "").Replace(" CO", "").Replace(" INCORPORATED", "").Replace(" INC", "").Replace(" ENTERPRISES", "").Replace(" ENTERPRISE", "").Replace(" LLC", "");

                                        string NameFinalOCR = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName.ToString().ToUpper();
                                        NameFinalOCR = Module1.func_RemoveSpecialWordsFromName(NameFinalOCR); //NameFinalOCR.Replace(" COMPANY", "").Replace(" CO", "").Replace(" INCORPORATED", "").Replace(" INC", "").Replace(" ENTERPRISES", "").Replace(" ENTERPRISE", "").Replace(" LLC", "");

                                        objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(NameFinalDB.ToString()), Module1.func_RemoveSpecialCharacter(NameFinalOCR));

                                        if (objRet50.PercentageMatch >= 90)
                                        {
                                            IsMatchNameConsignee = "Y";
                                        }
                                        if (IsMatchAddress == "Y") // && IsMatchNameConsignee == "Y"
                                        {
                                            strLogicConsignee = "YES";
                                        }
                                        else
                                        {
                                            strLogicConsignee = "NO";
                                        }

                                    }
                                    else
                                    {
                                        strLogicConsignee = "NO";
                                    }
                                }
                                else
                                {
                                    strLogicConsignee = "NA";
                                }
                            }
                           // Assgine to GP
                        clsModule.OCR_Matched_MasterDB_KeyConsignee = strLogicConsignee;
                        objStructF27.ContentsOfField = strLogicConsignee;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";

                        break;

                    case "C122": //  Consignee Details OCR_matched_Search enagine
                        string LogicCellConsignee = clsModule.OCR_Matched_MasterDB_KeyConsignee; //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(12, GridData1));
                        string strlogic1Consignee = string.Empty;
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "@@@@@")
                            {
                                if (!string.IsNullOrEmpty(LogicCellConsignee))
                                {
                                    if (LogicCellConsignee == "YES")
                                    {
                                        strlogic1Consignee = "NA";
                                    }
                                    else
                                    {
                                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeName) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeName != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeName != "@@@@@")
                                        {
                                            // TO DO => Need to use OCR_matched_Search enagine
                                            DataTable dtSearchEngine = new DataTable();
                                            string strflag = "C";
                                            dtSearchEngine = GetDataFromSearchEngine(strflag);
                                            if (dtSearchEngine != null && dtSearchEngine.Rows.Count > 0)
                                            {
                                                strlogic1Consignee = "YES";
                                                if (dtSearchEngine != null)
                                                {
                                                    dtSearchEngine.Dispose();
                                                    dtSearchEngine = null;
                                                }
                                            }
                                            else
                                            {
                                                strlogic1Consignee = "NO";
                                            }
                                        }
                                        else
                                        {
                                            strlogic1Consignee = "NO";
                                        }
                                    }
                                }
                            }
                            else
                            {
                                strlogic1Consignee = "NA";
                            }
                        clsModule.OCR_matched_SearchenagineConsignee = strlogic1Consignee;
                        //strlogic1Consignee = strlogic1Consignee.PadLeft(strlogic1Consignee.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                        objStructF27.ContentsOfField = strlogic1Consignee;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C0077A": // Consignee_Code# Final
                        string ConsigneeCodeFinal = string.Empty;
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "@@@@@")
                            {
                                string ConsigneeDetailsMasterDBKey = clsModule.OCR_Matched_MasterDB_KeyConsignee;
                                string ConsigneeDetailsSearchEngine = clsModule.OCR_matched_SearchenagineConsignee;
                                // Pickup => FRP001
                                if (!string.IsNullOrEmpty(ConsigneeDetailsMasterDBKey) && ConsigneeDetailsMasterDBKey == "YES")
                                {
                                    // STEP 1 - Check for HAT - Bill
                                    if (clsModule.clsPickupValuesFRP001HAT == "Y" && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal))
                                    {
                                        ConsigneeCodeFinal = clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal;
                                    }
                                    else
                                    {
                                        ConsigneeCodeFinal = clsModule.clsMasterValuesAPR001ConsigneeCode;
                                    }
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "5F";
                                }
                                // Master => ARP001
                                else if (!string.IsNullOrEmpty(ConsigneeDetailsSearchEngine) && ConsigneeDetailsSearchEngine == "YES")
                                {
                                    if (ClsSearchEngin.ClsSearchCommonStatusFlagConsignee == "F") // // Valid Code Found = GREEN 
                                    {
                                        ConsigneeCodeFinal = ClsSearchEngin.ClsSearchEnginConsigneeCode;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                    }
                                    else if (ClsSearchEngin.ClsSearchCommonStatusFlagConsignee == "AD" || ClsSearchEngin.ClsSearchCommonStatusFlagConsignee == "AO")  // Alias Code Found = GREEN
                                    {
                                        ConsigneeCodeFinal = ClsSearchEngin.ClsSearchEnginConsigneeCode;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                    }
                                    else if (ClsSearchEngin.ClsSearchCommonStatusFlagConsignee == "AF") // only address match then display name as OBL and code will be red only (ARP001 found one record and No alias data found)
                                    {
                                        ConsigneeCodeFinal = ClsSearchEngin.ClsSearchEnginConsigneeCode;
                                        objStructF27.Flag = "1"; //"0";
                                        objStructF27.Remarks = "6F";
                                    }
                                    else // Multiple code found = RED
                                    {
                                        ConsigneeCodeFinal = ClsSearchEngin.ClsSearchEnginConsigneeCode;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "6F";
                                    }
                                }
                                // OCR Value
                                else
                                {
                                    // TO DO => OCR Value
                                    ConsigneeCodeFinal = clsModule.clsPickupValuesFRP001ConsigneeCode; //"*****"; //clsModule.clsPickupValuesFRP001ConsigneeCode; 
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "1F";
                                }
                            }
                            else
                            {
                                clsModule.isConsigneeValidFlag = "F";
                                clsModule.isConsigneeAllGreenFlag = "F";
                                clsModule.isConsigneeCodeRedFlag = "T";
                            }

                            if (!string.IsNullOrEmpty(ConsigneeCodeFinal) && ConsigneeCodeFinal != "*****" && ConsigneeCodeFinal != "@@@@@")
                            {
                                ConsigneeCodeFinal = ConsigneeCodeFinal.PadLeft(ConsigneeCodeFinal.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                                objStructF27.ContentsOfField = ConsigneeCodeFinal;
                            }
                            else
                            {
                                objStructF27.ContentsOfField = "*****";
                                objStructF27.Flag = "0"; // added on 10-March-2021
                            }
                            clsModule.ConsigneeCodeFinal = ConsigneeCodeFinal;
                            objStructF27.X1 = 10;
                            objStructF27.Y1 = 10;
                            objStructF27.X2 = 20;
                            objStructF27.Y2 = 20;
                            objStructF27.PageNo = 998;
                            objStructF27.Status = "S";
                            // check code is valid or not
                            if (objStructF27.Flag == "0" && clsModule.isConsigneeValidFlag == "T")
                            {
                                clsModule.isConsigneeValidFlag = "F";
                            }
                            // Miscellaneous => check only consignee code is red 
                            if (objStructF27.Flag == "0")
                            {
                                clsModule.isConsigneeCodeRedFlag = "T";
                            }
                        break;

                    case "C0077": // Consignee Name Final
                        string ConsigneeNameFinal = string.Empty;
                            //string DBConsigneeNameFinal = string.Empty;
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "@@@@@")
                            {
                                string ConsigneeDetailsMasterDBKeyName = clsModule.OCR_Matched_MasterDB_KeyConsignee;
                                string ConsigneeDetailsSearchEngineName = clsModule.OCR_matched_SearchenagineConsignee;
                                string ConsigneeDetailsMaster = clsModule.clsMasterValuesAPR001ConsigneeDetails;
                                if (!string.IsNullOrEmpty(ConsigneeDetailsMasterDBKeyName) && ConsigneeDetailsMasterDBKeyName == "YES")
                                {
                                    // STEP 1 => Check for HAT
                                    if (clsModule.clsPickupValuesFRP001HAT == "Y")
                                    {
                                        // if valid Consignee code found in the master against consignee HAT bill
                                        if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATDetails))
                                        {
                                            string[] arrshipperDetilsMasterRAT = clsModule.clsMasterValuesAPR001ConsigneeHATDetails.Split('~');
                                            if (!string.IsNullOrEmpty(arrshipperDetilsMasterRAT[0]))
                                            {
                                                objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(arrshipperDetilsMasterRAT[0].ToString()), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesConsigneeName));
                                                if (objRet50.PercentageMatch >= 90)
                                                {
                                                    ConsigneeNameFinal = arrshipperDetilsMasterRAT[0].ToString();
                                                    objStructF27.Flag = "1";
                                                    objStructF27.Remarks = "5F";
                                                }
                                                else
                                                {
                                                    //if (clsF3.ClsOCRValues.ConsigneeNameOCRWithHighConfidence == "Y")
                                                    //{
                                                    //    ConsigneeNameFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName;
                                                    //    objStructF27.Flag = "1";
                                                    //    objStructF27.Remarks = "1F";
                                                    //}
                                                    //else
                                                    //{
                                                    ConsigneeNameFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName;
                                                    objStructF27.Flag = "0";
                                                    objStructF27.Remarks = "1F";
                                                    //}
                                                }
                                            }

                                        } // if NOT found for HAT
                                        else
                                        {
                                            //if (clsF3.ClsOCRValues.ConsigneeNameOCRWithHighConfidence == "Y")
                                            //{
                                            //    ConsigneeNameFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName;
                                            //    objStructF27.Flag = "1";
                                            //    objStructF27.Remarks = "1F";
                                            //}
                                            //else
                                            //{
                                            ConsigneeNameFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                            //}
                                        }
                                    }
                                    // Check for Normal case
                                    else
                                    {

                                        string[] arrshipperDetilsMasterCell = ConsigneeDetailsMaster.Trim().Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[0]))
                                        {
                                            string NameFinalDB = arrshipperDetilsMasterCell[0].ToString().ToUpper();
                                            string NameFinalDBOriginal = arrshipperDetilsMasterCell[0].ToString().ToUpper();
                                            NameFinalDB = Module1.func_RemoveSpecialWordsFromName(NameFinalDB); //NameFinalDB.Replace(" COMPANY", "").Replace(" CO", "").Replace(" INCORPORATED", "").Replace(" INC", "").Replace(" ENTERPRISES", "").Replace(" ENTERPRISE", "").Replace(" LLC", "");

                                            string NameFinalOCR = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName.ToString().ToUpper();
                                            string NameFinalOCROriginal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName.ToString().ToUpper();
                                            NameFinalOCR = Module1.func_RemoveSpecialWordsFromName(NameFinalOCR); //NameFinalOCR.Replace(" COMPANY", "").Replace(" CO", "").Replace(" INCORPORATED", "").Replace(" INC", "").Replace(" ENTERPRISES", "").Replace(" ENTERPRISE", "").Replace(" LLC", "");

                                            objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(NameFinalDB), Module1.func_RemoveSpecialCharacter(NameFinalOCR));
                                            if (objRet50.PercentageMatch > 90) // check the matching percentage of Name OBL and Name Master //2) if matched with master but Name OBL is blank then we can take it from master 
                                            {
                                                ConsigneeNameFinal = arrshipperDetilsMasterCell[0].ToString();
                                                objStructF27.Flag = "1";
                                                objStructF27.Remarks = "5F";
                                            }
                                            else
                                            {
                                                // check alias for name in master table if single record found then drop master name
                                                //string AddressLine1 = Regex.Split(clsModule.clsMasterValuesAPR001ConsigneeDetailsWithPipe, @"\|\|")[0];
                                                // DataTable dtAlias = GetAliasDataFromMaster(As400_ConnectionString, arrshipperDetilsMasterCell[3].ToString(), arrshipperDetilsMasterCell[4].ToString(), AddressLine1.ToString(), arrshipperDetilsMasterCell[2].ToString());
                                                string CustomerCode = string.Empty;
                                                if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeCode) && clsModule.clsMasterValuesAPR001ConsigneeCode != "*****")
                                                {
                                                    CustomerCode = "'" + clsModule.clsMasterValuesAPR001ConsigneeCode.ToString().Trim() + "'";
                                                    DataTable dtAlias = clsSearchCustomer.GetMatchedAliasFromARP033(CustomerCode, clsF3.ClsOCRValues.ClsOCRValuesConsigneeName);
                                                    if (dtAlias != null && dtAlias.Rows.Count == 1)
                                                    {
                                                        ConsigneeNameFinal = dtAlias.Rows[0]["NANAME"].ToString();//arrshipperDetilsMasterCell[0].ToString();
                                                        objStructF27.Flag = "1";
                                                        objStructF27.Remarks = "5F";
                                                    }
                                                    // check alias for name in master table if multiple record found then drop OBL name
                                                    else
                                                    {
                                                        string DBName = arrshipperDetilsMasterCell[0].Trim().ToUpper();
                                                        string OCRName = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName.Trim().ToUpper();
                                                        //string isWordFound = string.Empty;
                                                        int isWordFound = WordMatchFromName(DBName, OCRName);
                                                        if (isWordFound == 1)
                                                        {
                                                            ConsigneeNameFinal = arrshipperDetilsMasterCell[0].ToString();
                                                            objStructF27.Flag = "1";
                                                            objStructF27.Remarks = "5F";
                                                            // Above code commented and added below on 30-July-2021 to block one word logic
                                                            //objStructF27.Flag = "0";
                                                            //objStructF27.Remarks = "WF";
                                                        }
                                                        else if (isWordFound > 1)
                                                        {
                                                            ConsigneeNameFinal = arrshipperDetilsMasterCell[0].ToString();
                                                            objStructF27.Flag = "1";
                                                            objStructF27.Remarks = "5F";
                                                        }
                                                        else
                                                        {
                                                            // 1) Take orignal Names of DB and OCR (with out removing words) 2) Remove space 3) percentage match >=90 4) Shipper Name (Master DB) => Green with flag = 1 and remark = 5F ELSE drop OBL Name flag = 0 and remark = 1F 
                                                            objRet50 = ObjSbr.IQSBR050(NameFinalDBOriginal.ToString().Trim().Replace(" ", ""), NameFinalOCROriginal.ToString().Trim().Replace(" ", ""));
                                                            if (objRet50.PercentageMatch >= 90)
                                                            {
                                                                ConsigneeNameFinal = arrshipperDetilsMasterCell[0].ToString();
                                                                objStructF27.Flag = "1";
                                                                objStructF27.Remarks = "5F";
                                                            }
                                                            else
                                                            {
                                                                ConsigneeNameFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName;
                                                                objStructF27.Flag = "0";
                                                                objStructF27.Remarks = "1F";
                                                            }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    ConsigneeNameFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName;
                                                    objStructF27.Flag = "0";
                                                    objStructF27.Remarks = "1F";
                                                }
                                            }
                                        }
                                    }
                                }
                                else if (!string.IsNullOrEmpty(ConsigneeDetailsSearchEngineName) && ConsigneeDetailsSearchEngineName == "YES")
                                {
                                    if (ClsSearchEngin.ClsSearchCommonStatusFlagConsignee == "F") // Valid Code Found => Name = GREEN 
                                    {
                                        ConsigneeNameFinal = ClsSearchEngin.ClsSearchEnginConsigneeName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                    }
                                    else if (ClsSearchEngin.ClsSearchCommonStatusFlagConsignee == "AD")
                                    {
                                        ConsigneeNameFinal = ClsSearchEngin.ClsSearchEnginAliseConsigneeNameARP033;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                    }
                                    else if (ClsSearchEngin.ClsSearchCommonStatusFlagConsignee == "AO")  // Alias Code Found => Name = RED
                                    {

                                        string DBName = ClsSearchEngin.ClsSearchEnginConsigneeName.Trim().ToUpper();
                                        string OCRName = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName.Trim().ToUpper();
                                        //string isWordFound = string.Empty;
                                        int isWordFound = WordMatchFromName(DBName, OCRName);
                                        if (isWordFound == 1)
                                        {
                                            ConsigneeNameFinal = ClsSearchEngin.ClsSearchEnginConsigneeName;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "6F";
                                            // Above code commented and added below on 30-July-2021 to block one word logic
                                            //objStructF27.Flag = "0";
                                            //objStructF27.Remarks = "WF";
                                        }
                                        else if (isWordFound > 1)
                                        {
                                            ConsigneeNameFinal = ClsSearchEngin.ClsSearchEnginConsigneeName;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "6F";
                                        }
                                        else
                                        {
                                            // 1) Take orignal Names of DB and OCR (with out removing words) 2) Remove space 3) percentage match >=90 4) Shipper Name (Search Engine) => Green with flag = 1 and remark = 6F ELSE drop OBL Name flag = 0 and remark = 1F 
                                            objRet50 = ObjSbr.IQSBR050(DBName.ToString().Trim().Replace(" ", ""), OCRName.ToString().Trim().Replace(" ", ""));
                                            if (objRet50.PercentageMatch >= 90)
                                            {
                                                ConsigneeNameFinal = ClsSearchEngin.ClsSearchEnginConsigneeName;
                                                objStructF27.Flag = "1";
                                                objStructF27.Remarks = "6F";
                                            }
                                            else
                                            {
                                                ConsigneeNameFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName;
                                                objStructF27.Flag = "0";
                                                objStructF27.Remarks = "1F";
                                            }
                                        }


                                        //}
                                    }
                                    else if (ClsSearchEngin.ClsSearchCommonStatusFlagConsignee == "AF") // only address match then display name as OBL and code will be red only (ARP001 found one record and No alias data found)
                                    {

                                        string DBName = ClsSearchEngin.ClsSearchEnginConsigneeName.Trim().ToUpper();
                                        string OCRName = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName.Trim().ToUpper();
                                        //string isWordFound = string.Empty;
                                        int isWordFound = WordMatchFromName(DBName, OCRName);
                                        if (isWordFound == 1)
                                        {
                                            ConsigneeNameFinal = ClsSearchEngin.ClsSearchEnginConsigneeName;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "6F";
                                            // Above code commented and added below on 30-July-2021 to block one word logic
                                            //objStructF27.Flag = "0";
                                            //objStructF27.Remarks = "WF";
                                        }
                                        else if (isWordFound > 1)
                                        {
                                            ConsigneeNameFinal = ClsSearchEngin.ClsSearchEnginConsigneeName;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "6F";
                                        }
                                        else
                                        {
                                            // 1) Take orignal Names of DB and OCR (with out removing words) 2) Remove space 3) percentage match >=90 4) Shipper Name (Search Engine) => Green with flag = 1 and remark = 6F ELSE drop OBL Name flag = 0 and remark = 1F 
                                            objRet50 = ObjSbr.IQSBR050(DBName.ToString().Trim().Replace(" ", ""), OCRName.ToString().Trim().Replace(" ", ""));
                                            if (objRet50.PercentageMatch >= 90)
                                            {
                                                ConsigneeNameFinal = ClsSearchEngin.ClsSearchEnginConsigneeName;
                                                objStructF27.Flag = "1";
                                                objStructF27.Remarks = "6F";
                                            }
                                            else
                                            {
                                                ConsigneeNameFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName;
                                                objStructF27.Flag = "0";
                                                objStructF27.Remarks = "1F";
                                            }
                                        }
                                    }
                                    else // Multiple code found = RED
                                    {
                                        ConsigneeNameFinal = ClsSearchEngin.ClsSearchEnginConsigneeName;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "6F";
                                    }
                                }
                                else
                                {
                                    // If data is not found from master and serch enginee then match data with pick up and make it green
                                    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.clsPickupValuesFRP001ConsigneeName), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesConsigneeName));
                                    if (objRet50.PercentageMatch >= 90)
                                    {
                                        ConsigneeNameFinal = clsModule.clsPickupValuesFRP001ConsigneeName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "1F";
                                    }
                                    else
                                    {
                                        // check confidence level added on 18-March-2021
                                        // if (clsF3.ClsOCRValues.ConsigneeNameOCRWithHighConfidence == "Y") // Commented on 12-Aug-2021
                                        Regex regNotStartWithNumber = new Regex("^\\d[^<]+"); // Added on 12-Aug-2021 => OBL Name should not start with numeric
                                        if (clsF3.ClsOCRValues.ConsigneeNameOCRWithHighConfidence == "Y" && regNotStartWithNumber.IsMatch(clsF3.ClsOCRValues.ClsOCRValuesConsigneeName.Trim()) == false) // Added on 12-Aug-2021 => OBL Name should not start with numeric if it starts with numberic then make it as green
                                        {
                                            ConsigneeNameFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "1F";
                                        }
                                        else
                                        {
                                            ConsigneeNameFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                }
                            }
                            else
                            {
                                clsModule.isConsigneeValidFlag = "F";
                                clsModule.isConsigneeAllGreenFlag = "F";
                                clsModule.isConsigneeCodeRedFlag = "T";
                            }

                            if (!string.IsNullOrEmpty(ConsigneeNameFinal))
                            {
                                objStructF27.ContentsOfField = ConsigneeNameFinal;
                            }
                            else
                            {
                                objStructF27.ContentsOfField = "*****";
                                objStructF27.Flag = "0"; // added on 10-March-2021
                            }
                            // Added on 6-July-2021 => if shipper code and consignee code are same then make it Consignee Name RED so that it will not transfered to AS400 
                            if (clsModule.ShipperCodeFinal == clsModule.ConsigneeCodeFinal)
                            {
                                objStructF27.Flag = "0";
                            }
                            clsModule.clsFinalValuesConsigneeName = ConsigneeNameFinal;
                            objStructF27.X1 = 10;
                            objStructF27.Y1 = 10;
                            objStructF27.X2 = 20;
                            objStructF27.Y2 = 20;
                            objStructF27.PageNo = 998;
                            objStructF27.Status = "S";
                            // check code is valid or not
                            if (objStructF27.Flag == "0" && clsModule.isConsigneeValidFlag == "T")
                            {
                                clsModule.isConsigneeValidFlag = "F";
                            }
                            // Miscellaneous => check only consignee code is red 
                            if (objStructF27.Flag == "0" && clsModule.isConsigneeAllGreenFlag == "T")
                            {
                                clsModule.isConsigneeAllGreenFlag = "F";
                            }

                        break;
                    case "C0088": // Consignee address Line Final
                        string ConsigneeAddressLineFinal = string.Empty;
                        //string DBConsigneeAddressLineFinal = string.Empty;
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "@@@@@")
                            {
                                string ConsigneeDetailsMasterDBKeyAddressLine = clsModule.OCR_Matched_MasterDB_KeyConsignee;
                                string ConsigneeDetailsSearchEngineAddressLine = clsModule.OCR_matched_SearchenagineConsignee;
                                string ConsigneeDetailsMasterAddress = clsModule.clsMasterValuesAPR001ConsigneeDetails;
                                if (!string.IsNullOrEmpty(ConsigneeDetailsMasterDBKeyAddressLine) && ConsigneeDetailsMasterDBKeyAddressLine == "YES")
                                {
                                    // STEP 1 - Check for HAT
                                    if (clsModule.clsPickupValuesFRP001HAT == "Y")
                                    {
                                        // if valid shipper code found in the master against shipper RAT bill
                                        if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATDetails))
                                        {
                                            string[] arrshipperDetilsMasterRAT = clsModule.clsMasterValuesAPR001ConsigneeHATDetails.Split('~');
                                            if (!string.IsNullOrEmpty(arrshipperDetilsMasterRAT[1]))
                                            {
                                                ConsigneeAddressLineFinal = clsModule.clsMasterValuesAPR001ConsigneeDetailsWithPipe; //arrshipperDetilsMasterRAT[1].ToString();
                                                objStructF27.Flag = "1";
                                                objStructF27.Remarks = "5F";
                                            }

                                        }
                                        else
                                        {
                                            ConsigneeAddressLineFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                    // STEP 2 - Normal case
                                    else
                                    {
                                        string[] arrshipperDetilsMasterCell = ConsigneeDetailsMasterAddress.Trim().Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[1]))
                                        {
                                            ConsigneeAddressLineFinal = clsModule.clsMasterValuesAPR001ConsigneeDetailsWithPipe; //arrshipperDetilsMasterCell[1].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                }
                                else if (!string.IsNullOrEmpty(ConsigneeDetailsSearchEngineAddressLine) && ConsigneeDetailsSearchEngineAddressLine == "YES")

                                {
                                    ConsigneeAddressLineFinal = ClsSearchEngin.ClsSearchEnginConsigneeDetailsWithPipe; //ClsSearchEngin.ClsSearchEnginConsigneeAddressLine1; 
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else
                                {
                                    // If data is not found from master and serch enginee then match data with pick up and make it green
                                    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.clsPickupValuesFRP001ConsigneeAddress1), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1));
                                    if (objRet50.PercentageMatch >= 70)
                                    {
                                        ConsigneeAddressLineFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "1F";
                                    }
                                    else
                                    {
                                        // check confidence level 
                                        if (clsF3.ClsOCRValues.ConsigneeAddressLine1OCRWithHighConfidence == "Y")
                                        {
                                            ConsigneeAddressLineFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "1F";
                                        }
                                        else
                                        {
                                            ConsigneeAddressLineFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                }
                            }
                            else
                            {
                                clsModule.isConsigneeValidFlag = "F";
                                clsModule.isConsigneeAllGreenFlag = "F";
                                clsModule.isConsigneeCodeRedFlag = "T";
                            }

                            if (!string.IsNullOrEmpty(ConsigneeAddressLineFinal))
                            {

                                objStructF27.ContentsOfField = ConsigneeAddressLineFinal;
                            }
                            else
                            {
                                objStructF27.ContentsOfField = "*****";
                                objStructF27.Flag = "0"; // added on 10-March-2021
                            }
                            clsModule.clsFinalValuesConsigneeAddress1 = ConsigneeAddressLineFinal;
                            objStructF27.X1 = 10;
                            objStructF27.Y1 = 10;
                            objStructF27.X2 = 20;
                            objStructF27.Y2 = 20;
                            objStructF27.PageNo = 998;
                            objStructF27.Status = "S";
                            // check code is valid or not
                            if (objStructF27.Flag == "0" && clsModule.isConsigneeValidFlag == "T")
                            {
                                clsModule.isConsigneeValidFlag = "F";
                            }
                            // Miscellaneous => check only consignee code is red 
                            if (objStructF27.Flag == "0" && clsModule.isConsigneeAllGreenFlag == "T")
                            {
                                clsModule.isConsigneeAllGreenFlag = "F";
                            }
                       
                        break;

                    case "C088A": // Consignee City Final
                        string ConsigneeCityFinal = string.Empty;
                        //string DBConsigneeCityFinal = string.Empty;
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "@@@@@")
                            {
                                string ConsigneeDetailsMasterDBKeyCity = clsModule.OCR_Matched_MasterDB_KeyConsignee;
                                string ConsigneeDetailsSearchEngineCity = clsModule.OCR_matched_SearchenagineConsignee;
                                string ConsigneeDetailsMasterCity = clsModule.clsMasterValuesAPR001ConsigneeDetails;
                                if (!string.IsNullOrEmpty(ConsigneeDetailsMasterDBKeyCity) && ConsigneeDetailsMasterDBKeyCity == "YES")
                                {
                                    // STEP 1 - check shipper => HAT
                                    if (clsModule.clsPickupValuesFRP001HAT == "Y")
                                    {
                                        // if valid shipper code found in the master against shipper RAT bill
                                        if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATDetails))
                                        {
                                            string[] arrshipperDetilsMasterRAT = clsModule.clsMasterValuesAPR001ConsigneeHATDetails.Split('~');
                                            if (!string.IsNullOrEmpty(arrshipperDetilsMasterRAT[2]))
                                            {
                                                ConsigneeCityFinal = arrshipperDetilsMasterRAT[2].ToString();
                                                objStructF27.Flag = "1";
                                                objStructF27.Remarks = "5F";
                                            }

                                        }
                                        else
                                        {
                                            ConsigneeCityFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeCity;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                    else
                                    {
                                        string[] arrshipperDetilsMasterCell = ConsigneeDetailsMasterCity.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[2]))
                                        {
                                            ConsigneeCityFinal = arrshipperDetilsMasterCell[2].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                }
                                else if (!string.IsNullOrEmpty(ConsigneeDetailsSearchEngineCity) && ConsigneeDetailsSearchEngineCity == "YES")
                                {
                                    ConsigneeCityFinal = ClsSearchEngin.ClsSearchEnginConsigneeCity;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else
                                {
                                    // If data is not found from master and serch enginee then match data with pick up and make it green
                                    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.clsPickupValuesFRP001ConsigneeCity), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesConsigneeCity));
                                    if (objRet50.PercentageMatch == 100)
                                    {
                                        ConsigneeCityFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeCity;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "1F";
                                    }
                                    else
                                    {
                                        // check confidence level
                                        if (clsF3.ClsOCRValues.ConsigneeCityOCRWithHighConfidence == "Y")
                                        {
                                            ConsigneeCityFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeCity;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "1F";
                                        }
                                        else
                                        {
                                            ConsigneeCityFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeCity;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }

                                }
                            }
                            else
                            {
                                clsModule.isConsigneeValidFlag = "F";
                                clsModule.isConsigneeAllGreenFlag = "F";
                                clsModule.isConsigneeCodeRedFlag = "T";
                            }
                            if (!string.IsNullOrEmpty(ConsigneeCityFinal))
                            {

                                objStructF27.ContentsOfField = ConsigneeCityFinal;
                            }
                            else
                            {
                                objStructF27.ContentsOfField = "*****";
                                objStructF27.Flag = "0"; // added on 10-March-2021
                            }
                            clsModule.clsFinalValuesConsigneeCity = ConsigneeCityFinal;
                            objStructF27.X1 = 10;
                            objStructF27.Y1 = 10;
                            objStructF27.X2 = 20;
                            objStructF27.Y2 = 20;
                            objStructF27.PageNo = 998;
                            objStructF27.Status = "S";
                            // check code is valid or not
                            if (objStructF27.Flag == "0" && clsModule.isConsigneeValidFlag == "T")
                            {
                                clsModule.isConsigneeValidFlag = "F";
                            }
                            // Miscellaneous => check only consignee code is red 
                            if (objStructF27.Flag == "0" && clsModule.isConsigneeAllGreenFlag == "T")
                            {
                                clsModule.isConsigneeAllGreenFlag = "F";
                            }
                        
                        break;

                    case "C088B": // Consignee State Final
                        string ConsigneeStateFinal = string.Empty;
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "@@@@@")
                            {
                                string ConsigneeDetailsMasterDBKeyState = clsModule.OCR_Matched_MasterDB_KeyConsignee;//Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(12, GridDataShipperState));
                                string ConsigneeDetailsSearchEngineSate = clsModule.OCR_matched_SearchenagineConsignee;//Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(13, GridDataShipperState));
                                string ConsigneeDetailsMasterState = clsModule.clsMasterValuesAPR001ConsigneeDetails;//Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(11, GridDataShipperState));
                                if (!string.IsNullOrEmpty(ConsigneeDetailsMasterDBKeyState) && ConsigneeDetailsMasterDBKeyState == "YES")
                                {
                                    // STEP 1 -  check Consignee => HAT
                                    if (clsModule.clsPickupValuesFRP001HAT == "Y")
                                    {
                                        // if valid shipper code found in the master against shipper RAT bill
                                        if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATDetails))
                                        {
                                            string[] arrshipperDetilsMasterRAT = clsModule.clsMasterValuesAPR001ConsigneeHATDetails.Split('~');
                                            if (!string.IsNullOrEmpty(arrshipperDetilsMasterRAT[3]))
                                            {
                                                ConsigneeStateFinal = arrshipperDetilsMasterRAT[3].ToString();
                                                objStructF27.Flag = "1";
                                                objStructF27.Remarks = "5F";
                                            }
                                        }
                                        else
                                        {
                                            ConsigneeStateFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeState;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                    // Normal case
                                    else
                                    {

                                        string[] arrshipperDetilsMasterCell = ConsigneeDetailsMasterState.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[3]))
                                        {
                                            ConsigneeStateFinal = arrshipperDetilsMasterCell[3].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                }
                                else if (!string.IsNullOrEmpty(ConsigneeDetailsSearchEngineSate) && ConsigneeDetailsSearchEngineSate == "YES")
                                {
                                    ConsigneeStateFinal = ClsSearchEngin.ClsSearchEnginConsigneeState;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else
                                {
                                    // If data is not found from master and serch enginee then match data with pick up and make it green
                                    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.clsPickupValuesFRP001ConsigneeState), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesConsigneeState));
                                    if (objRet50.PercentageMatch == 100)
                                    {
                                        ConsigneeStateFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeState;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "1F";
                                    }
                                    else
                                    {
                                        if (clsF3.ClsOCRValues.ConsigneeStateOCRWithHighConfidence == "Y")
                                        {
                                            ConsigneeStateFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeState;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "1F";
                                        }
                                        else
                                        {
                                            ConsigneeStateFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeState;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }

                                }
                            }
                            else
                            {
                                clsModule.isConsigneeValidFlag = "F";
                                clsModule.isConsigneeAllGreenFlag = "F";
                                clsModule.isConsigneeCodeRedFlag = "T";

                            }
                            if (!string.IsNullOrEmpty(ConsigneeStateFinal))
                            {

                                objStructF27.ContentsOfField = ConsigneeStateFinal;
                            }
                            else
                            {
                                objStructF27.ContentsOfField = "*****";
                                objStructF27.Flag = "0"; // added on 10-March-2021
                            }
                            clsModule.clsFinalValuesConsigneeState = ConsigneeStateFinal;
                            objStructF27.X1 = 10;
                            objStructF27.Y1 = 10;
                            objStructF27.X2 = 20;
                            objStructF27.Y2 = 20;
                            objStructF27.PageNo = 998;
                            objStructF27.Status = "S";
                            // check code is valid or not
                            if (objStructF27.Flag == "0" && clsModule.isConsigneeValidFlag == "T")
                            {
                                clsModule.isConsigneeValidFlag = "F";
                            }
                            // Miscellaneous => check only consignee code is red 
                            if (objStructF27.Flag == "0" && clsModule.isConsigneeAllGreenFlag == "T")
                            {
                                clsModule.isConsigneeAllGreenFlag = "F";
                            }
                        
                        break;

                    case "C088C": // Consignee Zip Final
                        string ConsigneeZipFinal = string.Empty;
                        //string DBConsigneeZipFinal = string.Empty;
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "@@@@@")
                            {
                                string ConsigneeDetailsMasterDBKeyZip = clsModule.OCR_Matched_MasterDB_KeyConsignee;
                                string ConsigneeDetailsSearchEngineZip = clsModule.OCR_matched_SearchenagineConsignee;
                                string ConsigneeDetailsMasterZip = clsModule.clsMasterValuesAPR001ConsigneeDetails;
                                if (!string.IsNullOrEmpty(ConsigneeDetailsMasterDBKeyZip) && ConsigneeDetailsMasterDBKeyZip == "YES")
                                {
                                    if (clsModule.clsPickupValuesFRP001HAT == "Y")
                                    {
                                        // if valid shipper code found in the master against shipper RAT bill
                                        if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATDetails))
                                        {
                                            string[] arrshipperDetilsMasterRAT = clsModule.clsMasterValuesAPR001ConsigneeHATDetails.Split('~');
                                            if (!string.IsNullOrEmpty(arrshipperDetilsMasterRAT[4]))
                                            {
                                                ConsigneeZipFinal = arrshipperDetilsMasterRAT[4].ToString();
                                                objStructF27.Flag = "1";
                                                objStructF27.Remarks = "5F";
                                            }

                                        }
                                        else
                                        {
                                            ConsigneeZipFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeZip;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                    // Check for Normal case
                                    else
                                    {

                                        string[] arrshipperDetilsMasterCell = ConsigneeDetailsMasterZip.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[4]))
                                        {
                                            ConsigneeZipFinal = arrshipperDetilsMasterCell[4].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                }
                                else if (!string.IsNullOrEmpty(ConsigneeDetailsSearchEngineZip) && ConsigneeDetailsSearchEngineZip == "YES")
                                {
                                    ConsigneeZipFinal = ClsSearchEngin.ClsSearchEnginConsigneeZip;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else
                                {
                                    // If data is not found from master and serch enginee then match data with pick up and make it green
                                    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(clsModule.clsPickupValuesFRP001ConsigneeZip), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesConsigneeZip));
                                    if (objRet50.PercentageMatch == 100)
                                    {
                                        ConsigneeZipFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeZip;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "1F";
                                    }
                                    else
                                    {
                                        if (clsF3.ClsOCRValues.ConsigneeZipOCRWithHighConfidence == "Y")
                                        {
                                            ConsigneeZipFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeZip;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "1F";
                                        }
                                        else
                                        {
                                            ConsigneeZipFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeZip;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }

                                }
                            }
                            else
                            {
                                clsModule.isConsigneeValidFlag = "F";
                                clsModule.isConsigneeAllGreenFlag = "F";
                                clsModule.isConsigneeCodeRedFlag = "T";
                            }
                            if (!string.IsNullOrEmpty(ConsigneeZipFinal))
                            {

                                objStructF27.ContentsOfField = ConsigneeZipFinal;
                            }
                            else
                            {
                                objStructF27.ContentsOfField = "*****";
                                objStructF27.Flag = "0"; // added on 10-March-2021
                            }
                            clsModule.clsFinalValuesConsigneeZip = ConsigneeZipFinal;
                            objStructF27.X1 = 10;
                            objStructF27.Y1 = 10;
                            objStructF27.X2 = 20;
                            objStructF27.Y2 = 20;
                            objStructF27.PageNo = 998;
                            objStructF27.Status = "S";
                            // check code is valid or not
                            if (objStructF27.Flag == "0" && clsModule.isConsigneeValidFlag == "T")
                            {
                                clsModule.isConsigneeValidFlag = "F";
                            }
                            // Miscellaneous => check only consignee code is red 
                            if (objStructF27.Flag == "0" && clsModule.isConsigneeAllGreenFlag == "T")
                            {
                                clsModule.isConsigneeAllGreenFlag = "F";
                            }
                        
                        break;

                    case "C088D": // consignee Telephone Number Final
                        string ConsigneeTelFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim() != "@@@@@")
                        {
                            string ConsigneeDetailsMasterDBKeyTelePhoneNo = clsModule.OCR_Matched_MasterDB_KeyConsignee;
                            string ConsigneeDetailsSearchEngineTelePhoneNo = clsModule.OCR_matched_SearchenagineConsignee;
                            // Master Engine
                            if (!string.IsNullOrEmpty(ConsigneeDetailsMasterDBKeyTelePhoneNo) && ConsigneeDetailsMasterDBKeyTelePhoneNo == "YES")
                            {

                                if (clsModule.clsPickupValuesFRP001HAT == "Y")
                                {
                                    // if valid shipper code found in the master against shipper RAT bill
                                    if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal) && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeHATDetails))
                                    {
                                        objRet58 = ObjSbr.IQSBR058(Module1.func_RemoveSpecialCharacter(clsModule.clsMasterValuesAPR001ConsigneeHATTelePhoneNo), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo), true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
                                        if (objRet58.CharChanged <= 1)
                                        {
                                            ConsigneeTelFinal = clsModule.clsMasterValuesAPR001ConsigneeHATTelePhoneNo;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                        else
                                        {
                                            if (clsF3.ClsOCRValues.ConsigneeTelephoneOCRWithHighConfidence == "Y" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo != "*****")
                                            {
                                                ConsigneeTelFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                                                objStructF27.Flag = "1";
                                                objStructF27.Remarks = "1F";
                                            }
                                            else
                                            {
                                                ConsigneeTelFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                                                objStructF27.Flag = "0";
                                                objStructF27.Remarks = "1F";
                                            }

                                        }

                                    }
                                    else
                                    {
                                        if (clsF3.ClsOCRValues.ConsigneeTelephoneOCRWithHighConfidence == "Y" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo != "*****")
                                        {
                                            ConsigneeTelFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "1F";
                                        }
                                        else
                                        {
                                            ConsigneeTelFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                }
                                else
                                {

                                    objRet58 = ObjSbr.IQSBR058(Module1.func_RemoveSpecialCharacter(clsModule.clsMasterValuesAPR001ConsigneeTelePhoneNo), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo), true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
                                    if (objRet58.CharChanged <= 1)
                                    {
                                        ConsigneeTelFinal = clsModule.clsMasterValuesAPR001ConsigneeTelePhoneNo;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                    else
                                    {
                                        if (clsF3.ClsOCRValues.ConsigneeTelephoneOCRWithHighConfidence == "Y" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo != "*****")
                                        {
                                            ConsigneeTelFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "1F";
                                        }
                                        else
                                        {
                                            ConsigneeTelFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                }

                            }
                            // Search Engine 
                            else if (!string.IsNullOrEmpty(ConsigneeDetailsSearchEngineTelePhoneNo) && ConsigneeDetailsSearchEngineTelePhoneNo == "YES")
                            {
                                objRet58 = ObjSbr.IQSBR058(Module1.func_RemoveSpecialCharacter(ClsSearchEngin.ClsSearchEnginConsigneeTelePhoneNo), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo), true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
                                if (objRet58.CharChanged <= 1)
                                {
                                    ConsigneeTelFinal = ClsSearchEngin.ClsSearchEnginConsigneeTelePhoneNo;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else
                                {
                                    if (clsF3.ClsOCRValues.ConsigneeTelephoneOCRWithHighConfidence == "Y" && clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo != "*****")
                                    {
                                        ConsigneeTelFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "1F";
                                    }
                                    else
                                    {
                                        ConsigneeTelFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }

                            }
                            // OCR
                            else
                            {
                                if (clsF3.ClsOCRValues.ConsigneeTelephoneOCRWithHighConfidence == "Y")
                                {
                                    ConsigneeTelFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "1F";
                                }
                                else
                                {
                                    ConsigneeTelFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "1F";
                                }
                            }
                        }
                        if (!string.IsNullOrEmpty(ConsigneeTelFinal) && ConsigneeTelFinal != "*****" && ConsigneeTelFinal != "@@@@@" && ConsigneeTelFinal != "0")
                        {

                            objStructF27.ContentsOfField = ConsigneeTelFinal;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo))
                            {
                                ConsigneeTelFinal = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                                objStructF27.ContentsOfField = ConsigneeTelFinal;
                                objStructF27.Flag = "0";
                                objStructF27.Remarks = "1F";
                            }
                            else
                            {
                                objStructF27.ContentsOfField = "*****";
                            }
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        // Check Miscellaneous code => Code is RED but Name,Address,city,state,zip are GREEN
                        if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET2_ConsigneeMiscellaneousFlag"))
                        {
                            iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET2_ConsigneeMiscellaneousFlag");
                        }
                        if (clsModule.isConsigneeCodeRedFlag == "T" && clsModule.isConsigneeAllGreenFlag == "T" && clsF3.ClsOCRValues.ClsOCRCorrectConsigneeTagged == "T")
                        {
                            //iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET2_ConsigneeMiscellaneousFlag", "T"); // commented and added below on 16-June-2021 => If consignee name contains any special characters then dont return miscellaneous code flag  
                            var regexItem = new Regex("^[a-zA-Z0-9 ]*$");
                                if (regexItem.IsMatch(clsModule.clsFinalValuesConsigneeName))
                                {
                                    iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET2_ConsigneeMiscellaneousFlag", "T");
                                }
                                else
                                {
                                    iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET2_ConsigneeMiscellaneousFlag", "F");
                                }
                        }
                        else
                        {
                            iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET2_ConsigneeMiscellaneousFlag", "F");
                        }
                        break;

                    case "C24": // Mail To attached to shipper
                        string MailToAttachedShipper = string.Empty;
                        // step 1 - if shipper/consignee code final = pickup code final then dont fire query on DB
                        // step 2 - if shipper/consignee code final != pick up code final (in case of misscellaneous code) then fire query on DB and get the mailing address for shipper and consignee
                        // check mail to attached to shipper in DB
                        clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail = string.Empty;
                        clsModule.clsMasterValuesAPR001MailToAttachedShipperWithPipe = string.Empty;
                        if (!string.IsNullOrEmpty(clsModule.ShipperCodeFinal) && clsModule.ShipperCodeFinal != "*****")
                        {
                            // if (clsModule.clsPickupValuesFRP001ShipperCode != clsModule.ShipperCodeFinal) 
                            // Validate code before checking attachment
                            if (clsModule.OCR_Matched_MasterDB_KeyShipper == "YES" || clsModule.OCR_matched_SearchenagineShipper == "YES")
                            {
                                GetMailToAttachedData(clsModule.ShipperCodeFinal, "S");
                                if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail))
                                {
                                    MailToAttachedShipper = clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail;

                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(MailToAttachedShipper) && MailToAttachedShipper != "*****" && MailToAttachedShipper != "@@@@@")
                        {
                            // ShipperCollect = ShipperCollect.PadLeft(ShipperCollect.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                            objStructF27.ContentsOfField = MailToAttachedShipper;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C25": // Mail To attached to consignee
                        string MailToAttachedConsignee = string.Empty;
                        // step 1 - if shipper/consignee code final = pickup code final then dont fire query on DB
                        // step 2 - if shipper/consignee code final != pick up code final (in case of misscellaneous code) then fire query on DB and get the mailing address for shipper and consignee
                        // check mail to attached to shipper in DB
                        // check mail to attached to shipper in DB
                        clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail = string.Empty;
                        clsModule.clsMasterValuesAPR001MailToAttachedConsigneeWithPipe = string.Empty;
                        if (!string.IsNullOrEmpty(clsModule.ConsigneeCodeFinal) && clsModule.ConsigneeCodeFinal != "*****")
                        {
                            //if (clsModule.clsPickupValuesFRP001ConsigneeCode != clsModule.ConsigneeCodeFinal)
                            // Validate the code before checking attachement
                            if (clsModule.OCR_Matched_MasterDB_KeyConsignee == "YES" || clsModule.OCR_matched_SearchenagineConsignee == "YES")
                            {
                                GetMailToAttachedData(clsModule.ConsigneeCodeFinal, "C");
                                if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail))
                                {
                                    MailToAttachedConsignee = clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail;
                                }
                            }
                        }


                        if (!string.IsNullOrEmpty(MailToAttachedConsignee) && MailToAttachedConsignee != "*****" && MailToAttachedConsignee != "@@@@@")
                        {
                            // ShipperCollect = ShipperCollect.PadLeft(ShipperCollect.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                            objStructF27.ContentsOfField = MailToAttachedConsignee;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C0111": //  BillTo Code# => Pickup FRP001 / BillTo MailTo Code / 3rd party attached code
                        string BillToCodePickup = string.Empty;
                        // Bill To Code Attached to shipper

                        clsModule.isCMBTPCMBTCpresentShipper = "Y";
                        clsModule.isCMBTPCMBTCpresentValueShipper = string.Empty;

                        clsModule.isCMBTPCMBTCpresent = "Y";
                        clsModule.isCMBTPCMBTCpresentValue = string.Empty;

                        if (!string.IsNullOrEmpty(clsModule.ShipperCodeFinal) && clsModule.ShipperCodeFinal != "*****")
                        {
                            //if (clsModule.clsPickupValuesFRP001ShipperCode != clsModule.ShipperCodeFinal)
                            if (clsModule.OCR_Matched_MasterDB_KeyShipper == "YES" || clsModule.OCR_matched_SearchenagineShipper == "YES")
                            {
                                GetBillToAttachedData(clsModule.ShipperCodeFinal, "S");
                            }
                        }
                        if (!string.IsNullOrEmpty(clsModule.ConsigneeCodeFinal) && clsModule.ConsigneeCodeFinal != "*****")
                        {
                            // if (clsModule.clsPickupValuesFRP001ConsigneeCode != clsModule.ConsigneeCodeFinal)
                            if (clsModule.OCR_Matched_MasterDB_KeyConsignee == "YES" || clsModule.OCR_matched_SearchenagineConsignee == "YES")
                            {
                                GetBillToAttachedData(clsModule.ConsigneeCodeFinal, "C");
                            }
                        }

                        BillToCodePickup = clsModule.isCMBTPCMBTCpresentValueShipper + "#" + clsModule.isCMBTPCMBTCpresentValue;
                        // Get BillToAttached Code details => using code of isCMBTPCMBTCpresentValueShipper and clsModule.isCMBTPCMBTCpresentValue
                        clsModule.ShipperPrepaidCodeDetail = string.Empty;
                        clsModule.clsMasterValuesAPR001BillToAttachedShipperWithPipe = string.Empty;
                        if (!string.IsNullOrEmpty(clsModule.isCMBTPCMBTCpresentValueShipper) && clsModule.isCMBTPCMBTCpresentValueShipper != "*****")
                        {
                            if (clsModule.OCR_Matched_MasterDB_KeyShipper == "YES" || clsModule.OCR_matched_SearchenagineShipper == "YES")
                            {
                                GetBillToAttachedDataDetail(clsModule.isCMBTPCMBTCpresentValueShipper, "S");
                            }
                        }
                        clsModule.ConsigneeCollectCodeDetail = string.Empty;
                        clsModule.clsMasterValuesAPR001BillToAttachedConsigneeWithPipe = string.Empty;
                        if (!string.IsNullOrEmpty(clsModule.isCMBTPCMBTCpresentValue) && clsModule.isCMBTPCMBTCpresentValue != "*****")
                        {
                            if (clsModule.OCR_Matched_MasterDB_KeyConsignee == "YES" || clsModule.OCR_matched_SearchenagineConsignee == "YES")
                            {
                                GetBillToAttachedDataDetail(clsModule.isCMBTPCMBTCpresentValue, "C");
                            }
                        }

                        if (!string.IsNullOrEmpty(BillToCodePickup))
                        {
                            BillToCodePickup = BillToCodePickup.PadLeft(BillToCodePickup.Length, '0');
                            objStructF27.ContentsOfField = BillToCodePickup;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        //objStructF27.ContentsOfField = shipperCodePickup;
                        objStructF27.Status = "S";
                        break;

                    case "C20": // Shipper Prepaid code detail // 3rd party - Prepaid Code Address
                        string ShipperPrepaid = string.Empty;
                        //if (!string.IsNullOrEmpty(clsModule.ShipperPrepaidCodeDetail))
                        //{
                        if (!string.IsNullOrEmpty(clsModule.ShipperPrepaidCodeDetail))
                        {
                            ShipperPrepaid = clsModule.ShipperPrepaidCodeDetail;

                        }
                        //}
                        if (!string.IsNullOrEmpty(ShipperPrepaid) && ShipperPrepaid != "*****" && ShipperPrepaid != "@@@@@")
                        {

                            objStructF27.ContentsOfField = ShipperPrepaid;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C21": // Shipper collect code detail // NOT IN USE
                        #region NOT IN USE
                        string ShipperCollect = string.Empty;
                        if (!string.IsNullOrEmpty(clsModule.ShipperCollectCodeDetail))
                        {
                            ShipperCollect = clsModule.ShipperCollectCodeDetail;

                        }
                        if (!string.IsNullOrEmpty(ShipperCollect) && ShipperCollect != "*****" && ShipperCollect != "@@@@@")
                        {
                            // ShipperCollect = ShipperCollect.PadLeft(ShipperCollect.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                            objStructF27.ContentsOfField = ShipperCollect;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        #endregion
                        break;
                    case "C22": // Consignee Prepaid code detail // NOT IN USE
                        #region NOT IN USE
                        string ConsigneePrepaid = string.Empty;
                        if (!string.IsNullOrEmpty(clsModule.ConsigneePrepaidCodeDetail))
                        {
                            ConsigneePrepaid = clsModule.ConsigneePrepaidCodeDetail;

                        }
                        if (!string.IsNullOrEmpty(ConsigneePrepaid) && ConsigneePrepaid != "*****" && ConsigneePrepaid != "@@@@@")
                        {
                            //ConsigneePrepaid = ConsigneePrepaid.PadLeft(ConsigneePrepaid.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                            objStructF27.ContentsOfField = ConsigneePrepaid;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        #endregion
                        break;

                    case "C23": // Consignee collect code detail // 3rd party - collect code address
                        string ConsigneeCollect = string.Empty;
                        if (!string.IsNullOrEmpty(clsModule.ConsigneeCollectCodeDetail))
                        {
                            ConsigneeCollect = clsModule.ConsigneeCollectCodeDetail;

                        }
                        //}
                        if (!string.IsNullOrEmpty(ConsigneeCollect) && ConsigneeCollect != "*****" && ConsigneeCollect != "@@@@@")
                        {

                            objStructF27.ContentsOfField = ConsigneeCollect;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C0222": // BillTo Details (~Sep) => Pickup => Name~Add~City~State~Zip~TelNo  =>> NOT IN USE
                        #region NOT IN USE
                        string BillToDetailsPickup = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {
                            if (clsModule.isCMBTPCMBTCpresent == "Y")
                            {
                                BillToDetailsPickup = "NA";
                            }
                            else
                            {
                                BillToDetailsPickup = "NA";
                            }
                            // Temp code commented -- Start
                            //if (!string.IsNullOrEmpty(clsModule.clsPickupValuesFRP001BillToCode))
                            //{
                            //    BillToDetailsPickup = clsModule.clsPickupValuesFRP001BillToName + "~" + clsModule.clsPickupValuesFRP001BillToAddress1 + "~" + clsModule.clsPickupValuesFRP001BillToCity + "~" + clsModule.clsPickupValuesFRP001BillToState + "~" + clsModule.clsPickupValuesFRP001BillToZip;
                            //}
                            // Temp code commented -- end

                        }
                        if (!string.IsNullOrEmpty(BillToDetailsPickup) && BillToDetailsPickup != "NA")
                        {
                            //BillToDetailsPickup = BillToDetailsPickup.PadLeft(BillToDetailsPickup.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                            objStructF27.ContentsOfField = BillToDetailsPickup; ;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        #endregion
                        break;

                    case "C0999": // BillTo Master KeyField =>  Code_Pickup/TelNo_OCR
                        #region NOT IN USE
                        //if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        //{
                        //    string strBillToCodePickup = string.Empty;
                        //    if (clsModule.isCMBTPCMBTCpresent == "Y")
                        //    {
                        //        strBillToCodePickup = "NA";
                        //    }
                        //    else
                        //    {
                        //        strBillToCodePickup = "NA";
                        //    }
                        #region Temp code commented 
                        //strBillToCodePickup = clsModule.clsPickupValuesFRP001BillToCode;
                        //if (!string.IsNullOrEmpty(strBillToCodePickup))
                        //{

                        //        string strBillToCodeMasterOrTelNoOCR = string.Empty;
                        //        string strBillToDetailsMaster = string.Empty;
                        //        string strflag = "B";
                        //        GetShipperDetailsMaster(strBillToCodePickup, strflag, ref strBillToCodeMasterOrTelNoOCR, ref strBillToDetailsMaster);
                        //        string BillToCodeMasterOrTelOCR = strBillToCodeMasterOrTelNoOCR;
                        //        if (!string.IsNullOrEmpty(BillToCodeMasterOrTelOCR))
                        //        {
                        //            BillToCodeMasterOrTelOCR = BillToCodeMasterOrTelOCR.PadLeft(BillToCodeMasterOrTelOCR.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                        //            objStructF27.ContentsOfField = BillToCodeMasterOrTelOCR;
                        //        }
                        //        else
                        //        {
                        //            objStructF27.ContentsOfField = "@@@@@";
                        //        }
                        //}
                        #endregion
                        //}
                        //else
                        //{
                        //    objStructF27.ContentsOfField = "NA";
                        //}

                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {
                            objStructF27.ContentsOfField = "@@@@@";
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }

                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        #endregion
                        break;

                    case "C1000": // BillTo Details (~Sep) =>  Master
                        #region NOT IN USE
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {
                            objStructF27.ContentsOfField = "@@@@@";
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }

                        //if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim()) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        //{
                        //    string strBillToCodePickupDet = string.Empty;
                        //    if (clsModule.isCMBTPCMBTCpresent == "Y")
                        //    {
                        //        strBillToCodePickupDet = "NA";
                        //    }
                        //    else
                        //    {
                        //        strBillToCodePickupDet = "NA";
                        //    }
                        #region Temp code commented 
                        //strBillToCodePickupDet = clsModule.clsPickupValuesFRP001BillToCode;
                        //if (!string.IsNullOrEmpty(strBillToCodePickupDet))
                        //{

                        //        string strBillToCodeMasterOrTelNoOCR = string.Empty;
                        //        string strBillToDetailsMaster = string.Empty;
                        //        string strflag = "B";
                        //        GetShipperDetailsMaster(strBillToCodePickupDet, strflag, ref strBillToCodeMasterOrTelNoOCR, ref strBillToDetailsMaster);
                        //        string BillToDetailMaster = strBillToDetailsMaster;
                        //        if (!string.IsNullOrEmpty(BillToDetailMaster))
                        //        {
                        //            BillToDetailMaster = BillToDetailMaster.PadLeft(BillToDetailMaster.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                        //            objStructF27.ContentsOfField = BillToDetailMaster;
                        //        }
                        //        else
                        //        {
                        //            objStructF27.ContentsOfField = "@@@@@";
                        //        }


                        //}
                        #endregion
                        //}

                        //else
                        //{
                        //    objStructF27.ContentsOfField = "NA";
                        //}
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        #endregion
                        break;

                    case "C1111": // BillTo Details OCR_Matched_MasterDB_Key => Yes/No
                        string strLogicBillTo = string.Empty;
                        clsModule.isBillToAttachedShipper = string.Empty;
                        clsModule.isBillToAttachedConsignee = string.Empty;
                        clsModule.isMailToAttachedShipper = string.Empty;
                        clsModule.isMailToAttachedConsignee = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {
                            #region "As discussed Code Commented on 5-March-2021"
                            //// check Maill to attached with shipper 
                            //if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail) && clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail != "*****")
                            //{
                            //    // Step 1 :- Check Maill to attached with shipper
                            //    bool isNameMatched = false;
                            //    bool isAddressMatched = false;
                            //    string MasterMailToAttachedShipperName = clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail.Split('~')[0];
                            //    string MasterMailToAttachedShipperAddress = clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail.Split('~')[1] + " " + clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail.Split('~')[2] + " " + clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail.Split('~')[3] + " " + clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail.Split('~')[4];
                            //    string OCRBillToAddress = clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1 + " " + clsF3.ClsOCRValues.ClsOCRValuesBillToCity + " " + clsF3.ClsOCRValues.ClsOCRValuesBillToState + " " + clsF3.ClsOCRValues.ClsOCRValuesBillToZip;
                            //    // check Name is matched or not
                            //    if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToName) && clsF3.ClsOCRValues.ClsOCRValuesBillToName != "*****")
                            //    {
                            //        //if (MasterMailToAttachedShipperName.Contains("%"))
                            //        if (clsF3.ClsOCRValues.ClsOCRValuesBillToName.Contains("%"))
                            //        {
                            //            string[] arrMasterMailToAttachedShipperName = MasterMailToAttachedShipperName.Split('%');
                            //            string[] arrOCRMailToAttachedShipperName = clsF3.ClsOCRValues.ClsOCRValuesBillToName.Split('%');
                            //            for (int i = 0; i < arrMasterMailToAttachedShipperName.Length; i++)
                            //            {
                            //                for (int j = 0; j < arrOCRMailToAttachedShipperName.Length; j++)
                            //                {
                            //                    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(arrMasterMailToAttachedShipperName[i]), Module1.func_RemoveSpecialCharacter(arrOCRMailToAttachedShipperName[j]));
                            //                    if (objRet50.PercentageMatch >= 90)
                            //                    {
                            //                        isNameMatched = true;
                            //                    }
                            //                }
                            //            }

                            //        }
                            //        if (isNameMatched == false && (clsF3.ClsOCRValues.ClsOCRValuesBillToName.Contains("C/O") || clsF3.ClsOCRValues.ClsOCRValuesBillToName.Contains("c/o") || clsF3.ClsOCRValues.ClsOCRValuesBillToName.Contains("/")))
                            //        {
                            //            string[] arrMasterMailToAttachedShipperName = MasterMailToAttachedShipperName.Split('/');
                            //            string[] arrOCRMailToAttachedShipperName = clsF3.ClsOCRValues.ClsOCRValuesBillToName.Split('/');
                            //            for (int i = 0; i < arrMasterMailToAttachedShipperName.Length; i++)
                            //            {
                            //                for (int j = 0; j < arrOCRMailToAttachedShipperName.Length; j++)
                            //                {
                            //                    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(arrMasterMailToAttachedShipperName[i]), Module1.func_RemoveSpecialCharacter(arrOCRMailToAttachedShipperName[j]));
                            //                    if (objRet50.PercentageMatch >= 90)
                            //                    {
                            //                        isNameMatched = true;
                            //                    }
                            //                }
                            //            }

                            //        }
                            //    }
                            //    // check Address is match or not
                            //    string IsMatchAddress = MatchAddress(OCRBillToAddress, MasterMailToAttachedShipperAddress, "B");
                            //    if (IsMatchAddress == "Y")
                            //    {
                            //        isAddressMatched = true;
                            //    }
                            //    if (isAddressMatched == true) // removed name condition
                            //    {
                            //        clsModule.isMailToAttachedShipper = "Y";
                            //    }
                            //    else
                            //    {
                            //        clsModule.isMailToAttachedShipper = "N";
                            //    }

                            //}
                            //// check Maill to attached with consignee 
                            //if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail) && clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail != "*****")
                            //{
                            //    // Step 1 :- Check Maill to attached with shipper
                            //    bool isNameMatched = false;
                            //    bool isAddressMatched = false;
                            //    string MasterMailToAttachedConsigneeName = clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail.Split('~')[0];
                            //    string MasterMailToAttachedConsigneeAddress = clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail.Split('~')[1] + " " + clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail.Split('~')[2] + " " + clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail.Split('~')[3] + " " + clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail.Split('~')[4];
                            //    string OCRBillToAddress = clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1 + " " + clsF3.ClsOCRValues.ClsOCRValuesBillToCity + " " + clsF3.ClsOCRValues.ClsOCRValuesBillToState + " " + clsF3.ClsOCRValues.ClsOCRValuesBillToZip;
                            //    // check Name is matched or not
                            //    if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToName) && clsF3.ClsOCRValues.ClsOCRValuesBillToName != "*****")
                            //    {
                            //        if (MasterMailToAttachedConsigneeName.Contains("%"))
                            //        {
                            //            string[] arrMasterMailToAttachedConsigneeName = MasterMailToAttachedConsigneeName.Split('%');
                            //            string[] arrOCRMailToAttachedConsigneeName = clsF3.ClsOCRValues.ClsOCRValuesBillToName.Split('%');
                            //            for (int i = 0; i < arrMasterMailToAttachedConsigneeName.Length; i++)
                            //            {
                            //                for (int j = 0; j < arrOCRMailToAttachedConsigneeName.Length; j++)
                            //                {
                            //                    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(arrMasterMailToAttachedConsigneeName[i]), Module1.func_RemoveSpecialCharacter(arrOCRMailToAttachedConsigneeName[j]));
                            //                    if (objRet50.PercentageMatch >= 90)
                            //                    {
                            //                        isNameMatched = true;
                            //                    }
                            //                }
                            //            }

                            //        }
                            //        else if (MasterMailToAttachedConsigneeName.Contains("c/o"))
                            //        {
                            //            string[] arrMasterMailToAttachedShipperName = MasterMailToAttachedConsigneeName.Split('/');
                            //            string[] arrOCRMailToAttachedShipperName = clsF3.ClsOCRValues.ClsOCRValuesBillToName.Split('/');
                            //            for (int i = 0; i < arrMasterMailToAttachedShipperName.Length; i++)
                            //            {
                            //                for (int j = 0; j < arrOCRMailToAttachedShipperName.Length; j++)
                            //                {
                            //                    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(arrMasterMailToAttachedShipperName[i]), Module1.func_RemoveSpecialCharacter(arrOCRMailToAttachedShipperName[j]));
                            //                    if (objRet50.PercentageMatch >= 90)
                            //                    {
                            //                        isNameMatched = true;
                            //                    }
                            //                }
                            //            }

                            //        }
                            //    }
                            //    // check Address is match or not
                            //    string IsMatchAddress = MatchAddress(OCRBillToAddress, MasterMailToAttachedConsigneeAddress, "B");
                            //    if (IsMatchAddress == "Y")
                            //    {
                            //        isAddressMatched = true;
                            //    }
                            //    if (isAddressMatched == true)
                            //    {
                            //        clsModule.isMailToAttachedConsignee = "Y";
                            //    }
                            //    else
                            //    {
                            //        clsModule.isMailToAttachedConsignee = "N";
                            //    }

                            //}

                            //// if MailTo Attach with shipper/consignee then return YES 
                            //if (clsModule.isMailToAttachedShipper == "Y" || clsModule.isMailToAttachedConsignee == "Y")
                            //{
                            //    strLogicBillTo = "YES";
                            //}
                            //// Bill to attached with shipper/consignee Billtomail to code
                            //else
                            //{
                            //    //strLogicBillTo = "NO";
                            //    //if (clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail == "*****" && clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail != "*****")
                            //    if (!string.IsNullOrEmpty(clsModule.ShipperPrepaidCodeDetail) || !string.IsNullOrEmpty(clsModule.ConsigneeCollectCodeDetail)) // Added on 10-Feb-2021
                            //    {
                            //        #region "BillToAttached Code"
                            //        // STEP 1 :- Check 3rd party prepaid detail with SHIPPER OBL 
                            //        bool isShipperMatch = false;
                            //        bool isConsigneeMatch = false;
                            //        // check for 3rd party prepaid detail with OBL
                            //        if (!string.IsNullOrEmpty(clsModule.isCMBTPCMBTCpresentValueShipper) && !string.IsNullOrEmpty(clsModule.ShipperPrepaidCodeDetail))
                            //        {
                            //            string[] arrMasterDetailsCell = clsModule.ShipperPrepaidCodeDetail.Split('~');
                            //            string AddressLine1 = clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1; //clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1;
                            //            string City = clsF3.ClsOCRValues.ClsOCRValuesBillToCity;//clsF3.ClsOCRValues.ClsOCRValuesShipperCity;
                            //            string State = clsF3.ClsOCRValues.ClsOCRValuesBillToState; //clsF3.ClsOCRValues.ClsOCRValuesShipperState;
                            //            string Zip = clsF3.ClsOCRValues.ClsOCRValuesBillToZip;//clsF3.ClsOCRValues.ClsOCRValuesShipperZip;
                            //            string BillToDetilsMasterCellOCR = AddressLine1 + " " + City + " " + State + " " + Zip;
                            //            string BillToDetilsMasterCellMst = arrMasterDetailsCell[1] + " " + arrMasterDetailsCell[2] + " " + arrMasterDetailsCell[3] + " " + arrMasterDetailsCell[4];
                            //            string IsMatchAddress = MatchAddress(BillToDetilsMasterCellOCR, BillToDetilsMasterCellMst, "S");
                            //            if (IsMatchAddress == "Y")
                            //            {
                            //                isShipperMatch = true;
                            //                clsModule.isBillToAttachedShipper = "Y";
                            //            }
                            //            else
                            //            {
                            //                isShipperMatch = false;
                            //                clsModule.isBillToAttachedShipper = "N";
                            //            }
                            //        }
                            //        if (!isShipperMatch && !string.IsNullOrEmpty(clsModule.isCMBTPCMBTCpresentValue) && !string.IsNullOrEmpty(clsModule.ConsigneeCollectCodeDetail))
                            //        {
                            //            string[] arrMasterDetailsCell = clsModule.ConsigneeCollectCodeDetail.Split('~');
                            //            string AddressLine1 = clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1;
                            //            string City = clsF3.ClsOCRValues.ClsOCRValuesBillToCity;
                            //            string State = clsF3.ClsOCRValues.ClsOCRValuesBillToState;
                            //            string Zip = clsF3.ClsOCRValues.ClsOCRValuesBillToZip;
                            //            string BillToDetilsMasterCellOCR = AddressLine1 + " " + City + " " + State + " " + Zip;
                            //            string BillToDetilsMasterCellMst = arrMasterDetailsCell[1] + " " + arrMasterDetailsCell[2] + " " + arrMasterDetailsCell[3] + " " + arrMasterDetailsCell[4];
                            //            string IsMatchAddress = MatchAddress(BillToDetilsMasterCellOCR, BillToDetilsMasterCellMst, "S");
                            //            if (IsMatchAddress == "Y")
                            //            {
                            //                isConsigneeMatch = true;
                            //                clsModule.isBillToAttachedConsignee = "Y";
                            //            }
                            //            else
                            //            {
                            //                isConsigneeMatch = false;
                            //                clsModule.isBillToAttachedConsignee = "N";
                            //            }

                            //        }

                            //        if (isShipperMatch || isConsigneeMatch)
                            //        {
                            //            strLogicBillTo = "YES";
                            //        }
                            //        else
                            //        {
                            //            strLogicBillTo = "NO";
                            //        }
                            //        #endregion
                            //    }
                            //    else
                            //    {
                            //        strLogicBillTo = "NO";
                            //    }
                            //}
                            #endregion
                            #region "As discussed Code commented on 5-March-2021"
                            //// Logic :- 1) If only MailToAttach with shipper or consignee then take  
                            //// BillTo Code Final = Shipper/Consignee code 
                            //// BillTo Name Final = Shipper/Consignee Name 
                            //// BillToAddress , city, state , Zip Final = MailToAttached Details address (CMMAD1~CMMAD2~CMMCIT~CMMST~CMMZIP)
                            //// Logic :- 2) If  MailToAttach + BillToAttached (shipper/consignee) then check with BillToAttached details if we don't found code then take MailToAttached Details

                            //// Step 1 :- Check for BillToAttached details 
                            ////if (!string.IsNullOrEmpty(clsModule.ShipperPrepaidCodeDetail) || !string.IsNullOrEmpty(clsModule.ConsigneeCollectCodeDetail)) // Added on 10-Feb-2021
                            ////{
                            //#region "BillToAttach => Match address logic"
                            //bool isShipperMatch = false;
                            //bool isConsigneeMatch = false;
                            //// Check BillToAttach_Shp detail with BillTo OBL
                            //if (!string.IsNullOrEmpty(clsModule.isCMBTPCMBTCpresentValueShipper) && !string.IsNullOrEmpty(clsModule.ShipperPrepaidCodeDetail))
                            //{
                            //    string[] arrMasterDetailsCell = clsModule.ShipperPrepaidCodeDetail.Split('~');
                            //    string AddressLine1 = clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1; //clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1;
                            //    string City = clsF3.ClsOCRValues.ClsOCRValuesBillToCity;//clsF3.ClsOCRValues.ClsOCRValuesShipperCity;
                            //    string State = clsF3.ClsOCRValues.ClsOCRValuesBillToState; //clsF3.ClsOCRValues.ClsOCRValuesShipperState;
                            //    string Zip = clsF3.ClsOCRValues.ClsOCRValuesBillToZip;//clsF3.ClsOCRValues.ClsOCRValuesShipperZip;
                            //    string BillToDetilsMasterCellOCR = AddressLine1 + " " + City + " " + State + " " + Zip;
                            //    string BillToDetilsMasterCellMst = arrMasterDetailsCell[1] + " " + arrMasterDetailsCell[2] + " " + arrMasterDetailsCell[3] + " " + arrMasterDetailsCell[4];
                            //    string IsMatchAddress = MatchAddress(BillToDetilsMasterCellOCR, BillToDetilsMasterCellMst, "S");
                            //    if (IsMatchAddress == "Y")
                            //    {
                            //        isShipperMatch = true;
                            //        clsModule.isBillToAttachedShipper = "Y";
                            //    }
                            //    else
                            //    {
                            //        isShipperMatch = false;
                            //        clsModule.isBillToAttachedShipper = "N";
                            //    }
                            //}
                            //// Check BillToAttach_Con detail with BillTo OBL
                            //if (!isShipperMatch && !string.IsNullOrEmpty(clsModule.isCMBTPCMBTCpresentValue) && !string.IsNullOrEmpty(clsModule.ConsigneeCollectCodeDetail))
                            //{
                            //    string[] arrMasterDetailsCell = clsModule.ConsigneeCollectCodeDetail.Split('~');
                            //    string AddressLine1 = clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1;
                            //    string City = clsF3.ClsOCRValues.ClsOCRValuesBillToCity;
                            //    string State = clsF3.ClsOCRValues.ClsOCRValuesBillToState;
                            //    string Zip = clsF3.ClsOCRValues.ClsOCRValuesBillToZip;
                            //    string BillToDetilsMasterCellOCR = AddressLine1 + " " + City + " " + State + " " + Zip;
                            //    string BillToDetilsMasterCellMst = arrMasterDetailsCell[1] + " " + arrMasterDetailsCell[2] + " " + arrMasterDetailsCell[3] + " " + arrMasterDetailsCell[4];
                            //    string IsMatchAddress = MatchAddress(BillToDetilsMasterCellOCR, BillToDetilsMasterCellMst, "S");
                            //    if (IsMatchAddress == "Y")
                            //    {
                            //        isConsigneeMatch = true;
                            //        clsModule.isBillToAttachedConsignee = "Y";
                            //    }
                            //    else
                            //    {
                            //        isConsigneeMatch = false;
                            //        clsModule.isBillToAttachedConsignee = "N";
                            //    }

                            //}
                            //#endregion
                            //if (isShipperMatch || isConsigneeMatch)
                            //{
                            //    strLogicBillTo = "YES";
                            //}
                            //// Step 2 :- IF BillToAttached details are NOT matched THEN check for MaillToAttached details
                            //else
                            //{
                            //    // IF MaillToAttached with Both shipper and consignee THEN dont do anything => Return "NO"
                            //    if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail) && clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail != "*****"
                            //        && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail) && clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail != "*****")
                            //    {
                            //        // Dont Do anything 
                            //        strLogicBillTo = "NO";
                            //    }
                            //    // IF MaillToAttached_Shp details found THEN tagged MaillToAttached with Shipper 
                            //    else if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail) && clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail != "*****")
                            //    {
                            //        clsModule.isMailToAttachedShipper = "Y";
                            //    }
                            //    else if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail) && clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail != "*****")
                            //    {
                            //        clsModule.isMailToAttachedConsignee = "Y";
                            //    }
                            //    //else { }

                            //    // Check IF MaillToAttached_Shp OR MaillToAttached_Shp  details found then Return "YES"
                            //    if (clsModule.isMailToAttachedShipper == "Y" || clsModule.isMailToAttachedConsignee == "Y")
                            //    {
                            //        // check if isMailToAttachedShipper/isMailToAttachedConsignee then check shipper and consignee code if it is valid then return as YES if it is not valid code then send it for search engine 
                            //        if (clsModule.isMailToAttachedShipper == "Y" && (clsModule.OCR_Matched_MasterDB_KeyShipper == "YES" || clsModule.OCR_matched_SearchenagineShipper == "YES"))
                            //        {
                            //            strLogicBillTo = "YES";
                            //        }
                            //        else if (clsModule.isMailToAttachedConsignee == "Y" && (clsModule.OCR_Matched_MasterDB_KeyConsignee == "YES" || clsModule.OCR_matched_SearchenagineConsignee == "YES"))
                            //        {
                            //            strLogicBillTo = "YES";
                            //        }
                            //        else
                            //        {
                            //            strLogicBillTo = "NO";
                            //        }

                            //    }
                            //}
                            //#region "Step 3 is repeated"
                            ////}
                            //// Step 3 :- IF BillToAttached details are NOT Found THEN check for MaillToAttached details (Check only attachment dont compare address with OBL)
                            ////else
                            ////{
                            ////    // IF MaillToAttached with Both shipper and consignee THEN dont do anything
                            ////    if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail) && clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail != "*****"
                            ////        && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail) && clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail != "*****")
                            ////    {
                            ////        // Dont Do anything 
                            ////        strLogicBillTo = "NO";
                            ////    }
                            ////    // IF MaillToAttached_Shp details found THEN tagged MaillToAttached with Shipper 
                            ////    else if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail) && clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail != "*****")
                            ////    {
                            ////        clsModule.isMailToAttachedShipper = "Y";
                            ////    }
                            ////    else if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail) && clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail != "*****")
                            ////    {
                            ////        clsModule.isMailToAttachedConsignee = "Y";
                            ////    }
                            ////    //else { }

                            ////    // Check IF MaillToAttached_Shp OR MaillToAttached_Shp  details found then Return "YES"
                            ////    if (clsModule.isMailToAttachedShipper == "Y" || clsModule.isMailToAttachedConsignee == "Y")
                            ////    {
                            ////        // check if isMailToAttachedShipper/isMailToAttachedConsignee then check shipper and consignee code if it is valid then return as YES if it is not valid code then send it for search engine 
                            ////        if (clsModule.isMailToAttachedShipper == "Y" && (clsModule.OCR_Matched_MasterDB_KeyShipper == "YES" || clsModule.OCR_matched_SearchenagineShipper == "YES"))
                            ////        {
                            ////            strLogicBillTo = "YES";
                            ////        }
                            ////        else if (clsModule.isMailToAttachedConsignee == "Y" && (clsModule.OCR_Matched_MasterDB_KeyShipper == "YES" || clsModule.OCR_matched_SearchenagineShipper == "YES"))
                            ////        {
                            ////            strLogicBillTo = "YES";
                            ////        }
                            ////        else
                            ////        {
                            ////            strLogicBillTo = "NO";
                            ////        }
                            ////    }
                            ////}
                            //#endregion
                            //// Step 4 :- strLogicBillTo is empty then assign it as NO
                            //if (string.IsNullOrEmpty(strLogicBillTo))
                            //{
                            //    strLogicBillTo = "NO";
                            //}
                            #endregion

                            #region "As discussed code added on 24-March-2021"
                            bool isShipperMatch = false;
                            bool isConsigneeMatch = false;
                            // Step 1 => Terms must be present
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesTerms) && clsF3.ClsOCRValues.ClsOCRValuesTerms != "*****")
                            {
                                // Step 2 => if Terms are prepaid and shipper code is valid then pick  BillToAttach_Shp details
                                if (clsF3.ClsOCRValues.ClsOCRValuesTerms == "PPD" && clsModule.isShipperValidFlag != "F"
                                    && !string.IsNullOrEmpty(clsModule.ShipperPrepaidCodeDetail) && clsModule.ShipperPrepaidCodeDetail != "*****")
                                {
                                    isShipperMatch = true;
                                    clsModule.isBillToAttachedShipper = "Y";
                                }
                                // Step 3 => if Terms are Collect and consignee code is valid then pick  BillToAttach_Con details
                                if (clsF3.ClsOCRValues.ClsOCRValuesTerms == "COL" && clsModule.isConsigneeValidFlag != "F"
                                    && !string.IsNullOrEmpty(clsModule.ConsigneeCollectCodeDetail) && clsModule.ConsigneeCollectCodeDetail != "*****")
                                {
                                    isConsigneeMatch = true;
                                    clsModule.isBillToAttachedConsignee = "Y";
                                }
                                // check BillToAttach for shipper or consignee
                                if (isShipperMatch || isConsigneeMatch)
                                {
                                    strLogicBillTo = "YES";
                                }
                                else
                                {
                                    // Step 4 => if Terms are prepaid and shipper code is valid then pick  MailAttach_Shp details
                                    if (clsF3.ClsOCRValues.ClsOCRValuesTerms == "PPD" && clsModule.isShipperValidFlag != "F"
                                        && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail)
                                        && clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail != "*****")
                                    {
                                        clsModule.isMailToAttachedShipper = "Y";
                                    }
                                    // Step 5 => if Terms are prepaid and Consignee code is valid then pick  MailAttach_Con details
                                    if (clsF3.ClsOCRValues.ClsOCRValuesTerms == "COL" && clsModule.isConsigneeValidFlag != "F"
                                       && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail)
                                       && clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail != "*****")
                                    {
                                        clsModule.isMailToAttachedConsignee = "Y";
                                    }
                                    if (clsModule.isMailToAttachedShipper == "Y" || clsModule.isMailToAttachedConsignee == "Y")
                                    {
                                        strLogicBillTo = "YES";
                                    }
                                }
                                // if nothing found then return NO
                                if (string.IsNullOrEmpty(strLogicBillTo))
                                {
                                    strLogicBillTo = "NO";
                                }
                            }
                            else
                            {
                                strLogicBillTo = "NO"; //  send for serach engine
                            }
                            #endregion
                        }
                        else
                        {
                            strLogicBillTo = "*****";
                        }

                        // Assgine to GP
                        clsModule.OCR_Matched_MasterDB_KeyBillTo = strLogicBillTo;
                        objStructF27.ContentsOfField = strLogicBillTo;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";

                        break;

                    case "C1222": //  BillTo Details OCR_matched_Search enagine
                        string strlogic1BillTo = string.Empty;
                        strlogic1BillTo = "NA";
                        //if (strlogic1BillTo == "NA")
                        //{
                        //    strlogic1BillTo = "NA";
                        //}
                        //else
                        //{
                        string LogicCellBillTo = clsModule.OCR_Matched_MasterDB_KeyBillTo;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {
                            if (!string.IsNullOrEmpty(LogicCellBillTo))
                            {
                                if (LogicCellBillTo == "YES")
                                {
                                    strlogic1BillTo = "NA";
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToName) && clsF3.ClsOCRValues.ClsOCRValuesBillToName != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToName != "@@@@@")
                                    {
                                        // BillTo Search engine execute only when shipper code final and consignee code final is present
                                        //if (!string.IsNullOrEmpty(clsModule.ShipperCodeFinal) && clsModule.ShipperCodeFinal != "*****" && 
                                        //    !string.IsNullOrEmpty(clsModule.ConsigneeCodeFinal) && clsModule.ConsigneeCodeFinal != "*****") // commented on 24-March-2021
                                        // search engine => (both shipper and consignee must be green) and (Not Attached to any Shipper or Consignee)
                                        if (!string.IsNullOrEmpty(clsModule.isShipperValidFlag) && clsModule.isShipperValidFlag != "F" &&
                                            !string.IsNullOrEmpty(clsModule.isConsigneeValidFlag) && clsModule.isConsigneeValidFlag != "F"
                                            && (string.IsNullOrEmpty(clsModule.ShipperPrepaidCodeDetail) || clsModule.ShipperPrepaidCodeDetail == "*****")
                                            && (string.IsNullOrEmpty(clsModule.ConsigneeCollectCodeDetail) || clsModule.ConsigneeCollectCodeDetail == "*****")
                                            && (string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail) || clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail == "*****")
                                            && (string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail) || clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail == "*****")
                                            )
                                        {
                                            DataTable dtSearchEngine = new DataTable();
                                            string strflag = "B";
                                            dtSearchEngine = GetDataFromSearchEngine(strflag);
                                            if (dtSearchEngine != null && dtSearchEngine.Rows.Count > 0)
                                            {
                                                strlogic1BillTo = "YES";
                                                if (dtSearchEngine != null)
                                                {
                                                    dtSearchEngine.Dispose();
                                                    dtSearchEngine = null;
                                                }
                                            }
                                            else
                                            {
                                                strlogic1BillTo = "NO";
                                            }
                                        }
                                        else
                                        {
                                            strlogic1BillTo = "NA";
                                        }
                                    }
                                    else
                                    {
                                        strlogic1BillTo = "NA";
                                    }
                                }
                            }
                        }
                        else
                        {
                            strlogic1BillTo = "NA";
                        }
                        //}
                        clsModule.OCR_matched_SearchenagineBillTo = strlogic1BillTo;
                        objStructF27.ContentsOfField = strlogic1BillTo;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C00777A": // BillTo_Code# Final
                        string BillToCodeFinal = string.Empty;
                        //string DBBillToCodeFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {

                            string BillToDetailsMasterDBKey = clsModule.OCR_Matched_MasterDB_KeyBillTo;
                            string BillToDetailsSearchEngine = clsModule.OCR_matched_SearchenagineBillTo;
                            if (!string.IsNullOrEmpty(BillToDetailsMasterDBKey) && BillToDetailsMasterDBKey == "YES")
                            {
                                if (clsModule.isMailToAttachedShipper == "Y")
                                {
                                    BillToCodeFinal = clsModule.ShipperCodeFinal;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "5F";
                                }
                                else if (clsModule.isMailToAttachedConsignee == "Y")
                                {
                                    BillToCodeFinal = clsModule.ConsigneeCodeFinal;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "5F";
                                }
                                else
                                {
                                    if (BillToCodeFinal == "*****" || string.IsNullOrEmpty(BillToCodeFinal)) // if MailToAttached with shipper and consignee not found then check for BillToAttached with shipper and consignee
                                    {
                                        #region BillToMailTo Code
                                        if (clsModule.isBillToAttachedShipper == "Y")
                                        {
                                            BillToCodeFinal = clsModule.isCMBTPCMBTCpresentValueShipper;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                        else if (clsModule.isBillToAttachedConsignee == "Y")
                                        {
                                            BillToCodeFinal = clsModule.isCMBTPCMBTCpresentValue;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                        #endregion
                                    }
                                }
                            }
                            // Master => ARP001
                            else if (!string.IsNullOrEmpty(BillToDetailsSearchEngine) && BillToDetailsSearchEngine == "YES")
                            {
                                if (ClsSearchEngin.ClsSearchCommonStatusFlagBillTo == "F") // Valid Code Found = GREEN 
                                {
                                    BillToCodeFinal = ClsSearchEngin.ClsSearchEnginBillToCode;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else if (ClsSearchEngin.ClsSearchCommonStatusFlagBillTo == "AD" || ClsSearchEngin.ClsSearchCommonStatusFlagBillTo == "AO") // Alias Code Found = GREEN
                                {
                                    BillToCodeFinal = ClsSearchEngin.ClsSearchEnginBillToCode;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else if (ClsSearchEngin.ClsSearchCommonStatusFlagBillTo == "AF") // only address match then display name as OBL and code will be red only (ARP001 found one record and No alias data found)
                                {
                                    BillToCodeFinal = ClsSearchEngin.ClsSearchEnginBillToCode;
                                    objStructF27.Flag = "1";//"0";
                                    objStructF27.Remarks = "6F";
                                }
                                else // Multiple code found = RED
                                {
                                    BillToCodeFinal = ClsSearchEngin.ClsSearchEnginBillToCode;
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "6F";
                                }
                            }
                            // OCR Value
                            else
                            {
                                // TO DO => OCR Value => Temp code commented
                                //BillToCodeFinal = clsModule.clsPickupValuesFRP001BillToCode; 

                                BillToCodeFinal = "";
                                objStructF27.Flag = "0";
                                objStructF27.Remarks = "1F";

                            }


                        }
                        if (!string.IsNullOrEmpty(BillToCodeFinal) && BillToCodeFinal != "*****" && BillToCodeFinal != "@@@@@")
                        {
                            BillToCodeFinal = BillToCodeFinal.PadLeft(BillToCodeFinal.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                            objStructF27.ContentsOfField = BillToCodeFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C00777": // BillTo Name Final
                        string BillToNameFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {

                            string BillToDetailsMasterDBKeyName = clsModule.OCR_Matched_MasterDB_KeyBillTo; //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(12, GridDataShipperName));
                            string BillToDetailsSearchEngineName = clsModule.OCR_matched_SearchenagineBillTo; //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(13, GridDataShipperName));
                            string BillToDetailsMaster = clsModule.clsMasterValuesAPR001BillToDetails;  //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(11, GridDataShipperName));
                            if (!string.IsNullOrEmpty(BillToDetailsMasterDBKeyName) && BillToDetailsMasterDBKeyName == "YES")
                            {
                                if (clsModule.isMailToAttachedShipper == "Y")
                                {
                                    if (!string.IsNullOrEmpty(clsModule.clsFinalValuesShipperName))
                                    {
                                        BillToNameFinal = clsModule.clsFinalValuesShipperName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                                else if (clsModule.isMailToAttachedConsignee == "Y")
                                {
                                    if (!string.IsNullOrEmpty(clsModule.clsFinalValuesConsigneeName))
                                    {
                                        BillToNameFinal = clsModule.clsFinalValuesConsigneeName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                                else if (clsModule.isMailToAttachedShipper != "Y" && clsModule.isMailToAttachedConsignee != "Y")
                                {
                                    #region BillToMailTo Code
                                    if (clsModule.isBillToAttachedShipper == "Y")
                                    {
                                        string[] arrshipperDetilsMasterCell = clsModule.ShipperPrepaidCodeDetail.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[0]))
                                        {
                                            BillToNameFinal = arrshipperDetilsMasterCell[0].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                    else if (clsModule.isBillToAttachedConsignee == "Y")
                                    {
                                        string[] arrshipperDetilsMasterCell = clsModule.ConsigneeCollectCodeDetail.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[0]))
                                        {
                                            BillToNameFinal = arrshipperDetilsMasterCell[0].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                    // If Bill to code matches with Shipper and consignee then drop the OCR Bill To OCR
                                    #endregion
                                }
                                else
                                {
                                    if (clsF3.ClsOCRValues.BillToNameOCRWithHighConfidence == "Y")
                                    {
                                        BillToNameFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "1F";
                                    }
                                    else
                                    {
                                        BillToNameFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToName;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }

                                }

                            }
                            else if (!string.IsNullOrEmpty(BillToDetailsSearchEngineName) && BillToDetailsSearchEngineName == "YES")

                            {
                                #region Commented code
                                //BillToNameFinal = ClsSearchEngin.ClsSearchEnginBillToName;
                                //if (ClsSearchEngin.ClsSearchEnginAliseShipperName != "Y")
                                //{
                                //    BillToNameFinal = ClsSearchEngin.ClsSearchEnginBillToName;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                //else
                                //{ 
                                //objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(ClsSearchEngin.ClsSearchEnginBillToName), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesBillToName));
                                //if (objRet50.PercentageMatch >= 90)
                                //{
                                //    BillToNameFinal = ClsSearchEngin.ClsSearchEnginBillToName;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                //else
                                //{
                                //    BillToNameFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToName;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "6F";
                                //}
                                //}
                                #endregion
                                if (ClsSearchEngin.ClsSearchCommonStatusFlagBillTo == "F") // Valid Code Found => Name = GREEN 
                                {
                                    BillToNameFinal = ClsSearchEngin.ClsSearchEnginBillToName;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else if (ClsSearchEngin.ClsSearchCommonStatusFlagBillTo == "AD")
                                {
                                    BillToNameFinal = ClsSearchEngin.ClsSearchEnginAliseBillToNameARP033;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else if (ClsSearchEngin.ClsSearchCommonStatusFlagBillTo == "AO")  // Alias Code Found => Name = RED
                                {
                                    //if (clsF3.ClsOCRValues.BillToNameOCRWithHighConfidence == "Y")
                                    //{
                                    //    BillToNameFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToName;
                                    //    objStructF27.Flag = "1";
                                    //    objStructF27.Remarks = "1F";
                                    //}
                                    //else
                                    //{
                                    //BillToNameFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToName;
                                    //objStructF27.Flag = "0";
                                    //objStructF27.Remarks = "1F";
                                    //}

                                    string DBName = ClsSearchEngin.ClsSearchEnginBillToName.Trim().ToUpper();
                                    string OCRName = clsF3.ClsOCRValues.ClsOCRValuesBillToName.Trim().ToUpper();
                                    //string isWordFound = string.Empty;
                                    int isWordFound = WordMatchFromName(DBName, OCRName);
                                    if (isWordFound == 1)
                                    {
                                        BillToNameFinal = ClsSearchEngin.ClsSearchEnginBillToName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                        // Above code commented and added below on 30-July-2021 to block one word logic
                                        //objStructF27.Flag = "0";
                                        //objStructF27.Remarks = "WF";
                                    }
                                    else if (isWordFound > 1)
                                    {
                                        BillToNameFinal = ClsSearchEngin.ClsSearchEnginBillToName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                    }
                                    else
                                    {
                                        // 1) Take orignal Names of DB and OCR (with out removing words) 2) Remove space 3) percentage match >=90 4) Shipper Name (Search Engine) => Green with flag = 1 and remark = 6F ELSE drop OBL Name flag = 0 and remark = 1F 
                                        objRet50 = ObjSbr.IQSBR050(DBName.ToString().Trim().Replace(" ", ""), OCRName.ToString().Trim().Replace(" ", ""));
                                        if (objRet50.PercentageMatch >= 90)
                                        {
                                            BillToNameFinal = ClsSearchEngin.ClsSearchEnginBillToName;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "6F";
                                        }
                                        else
                                        {
                                            BillToNameFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToName;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                }
                                else if (ClsSearchEngin.ClsSearchCommonStatusFlagBillTo == "AF") // only address match then display name as OBL and code will be red only (ARP001 found one record and No alias data found)
                                {
                                    //if (clsF3.ClsOCRValues.BillToNameOCRWithHighConfidence == "Y")
                                    //{
                                    //    BillToNameFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToName;
                                    //    objStructF27.Flag = "1";
                                    //    objStructF27.Remarks = "1F";
                                    //}
                                    //else
                                    //{
                                    //BillToNameFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToName;
                                    //objStructF27.Flag = "0";
                                    //objStructF27.Remarks = "1F";
                                    //}
                                    string DBName = ClsSearchEngin.ClsSearchEnginBillToName.Trim().ToUpper();
                                    string OCRName = clsF3.ClsOCRValues.ClsOCRValuesBillToName.Trim().ToUpper();
                                    //string isWordFound = string.Empty;
                                    int isWordFound = WordMatchFromName(DBName, OCRName);
                                    if (isWordFound == 1)
                                    {
                                        BillToNameFinal = ClsSearchEngin.ClsSearchEnginBillToName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                        // Above code commented and added below on 30-July-2021 to block one word logic
                                        //objStructF27.Flag = "0";
                                        //objStructF27.Remarks = "WF";
                                    }
                                    else if (isWordFound > 1)
                                    {
                                        BillToNameFinal = ClsSearchEngin.ClsSearchEnginBillToName;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "6F";
                                    }
                                    else
                                    {
                                        // 1) Take orignal Names of DB and OCR (with out removing words) 2) Remove space 3) percentage match >=90 4) Shipper Name (Search Engine) => Green with flag = 1 and remark = 6F ELSE drop OBL Name flag = 0 and remark = 1F 
                                        objRet50 = ObjSbr.IQSBR050(DBName.ToString().Trim().Replace(" ", ""), OCRName.ToString().Trim().Replace(" ", ""));
                                        if (objRet50.PercentageMatch >= 90)
                                        {
                                            BillToNameFinal = ClsSearchEngin.ClsSearchEnginBillToName;
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "6F";
                                        }
                                        else
                                        {
                                            BillToNameFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToName;
                                            objStructF27.Flag = "0";
                                            objStructF27.Remarks = "1F";
                                        }
                                    }
                                }
                                else // Multiple code found => Name = RED
                                {
                                    BillToNameFinal = ClsSearchEngin.ClsSearchEnginBillToName;
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "6F";
                                }
                            }
                            else
                            {
                                // check the confidance level if it is greater than 90 then make it green
                                //if (clsF3.ClsOCRValues.BillToNameOCRWithHighConfidence == "Y")
                                //{
                                //    BillToNameFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToName;
                                //    objStructF27.Flag = "1";
                                //    objStructF27.Remarks = "1F";
                                //}
                                //else
                                //{
                                BillToNameFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToName;
                                objStructF27.Flag = "0";
                                objStructF27.Remarks = "1F";
                                //}
                            }

                        }
                        if (!string.IsNullOrEmpty(BillToNameFinal))
                        {

                            objStructF27.ContentsOfField = BillToNameFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        clsModule.clsFinalValuesBillToName = BillToNameFinal;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;
                    case "C00888": // BillTo address Line Final
                        string BillToAddressLineFinal = string.Empty;
                        //string DBBillToAddressLineFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {
                            //if (clsModule.isMisscellenaousBillToCodePickup == "Y")
                            //{
                            string BillToDetailsMasterDBKeyAddressLine = clsModule.OCR_Matched_MasterDB_KeyBillTo;
                            string BillToDetailsSearchEngineAddressLine = clsModule.OCR_matched_SearchenagineBillTo;
                            string BillToDetailsMasterAddress = clsModule.clsMasterValuesAPR001BillToDetails;
                            if (!string.IsNullOrEmpty(BillToDetailsMasterDBKeyAddressLine) && BillToDetailsMasterDBKeyAddressLine == "YES")
                            {

                                if (clsModule.isMailToAttachedShipper == "Y")
                                {
                                    if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperWithPipe))
                                    {
                                        BillToAddressLineFinal = clsModule.clsMasterValuesAPR001MailToAttachedShipperWithPipe;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                                else if (clsModule.isMailToAttachedConsignee == "Y")
                                {
                                    if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeWithPipe))
                                    {
                                        BillToAddressLineFinal = clsModule.clsMasterValuesAPR001MailToAttachedConsigneeWithPipe;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                                else if (clsModule.isMailToAttachedShipper != "Y" && clsModule.isMailToAttachedConsignee != "Y")
                                {
                                    #region BillToMailTo Code
                                    if (clsModule.isBillToAttachedShipper == "Y")
                                    {
                                        string[] arrshipperDetilsMasterCell = clsModule.ShipperPrepaidCodeDetail.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[1]))
                                        {
                                            BillToAddressLineFinal = clsModule.clsMasterValuesAPR001BillToAttachedShipperWithPipe; //arrshipperDetilsMasterCell[1].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                    else if (clsModule.isBillToAttachedConsignee == "Y")
                                    {
                                        string[] arrshipperDetilsMasterCell = clsModule.ConsigneeCollectCodeDetail.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[1]))
                                        {
                                            BillToAddressLineFinal = clsModule.clsMasterValuesAPR001BillToAttachedConsigneeWithPipe; //arrshipperDetilsMasterCell[1].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                    // If Bill to code matches with Shipper and consignee then drop the OCR Bill To OCR
                                    #endregion
                                }

                                else
                                {
                                    BillToAddressLineFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1;
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "1F";
                                }
                            }
                            else if (!string.IsNullOrEmpty(BillToDetailsSearchEngineAddressLine) && BillToDetailsSearchEngineAddressLine == "YES")

                            {
                                BillToAddressLineFinal = ClsSearchEngin.ClsSearchEnginBillToDetailsWithPipe; //ClsSearchEngin.ClsSearchEnginBillToAddressLine1;
                                objStructF27.Flag = "1";
                                objStructF27.Remarks = "6F";
                            }
                            else
                            {
                                // TO DO => OCR 
                                BillToAddressLineFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1;
                                objStructF27.Flag = "0";
                                objStructF27.Remarks = "1F";
                            }

                        }
                        if (!string.IsNullOrEmpty(BillToAddressLineFinal))
                        {
                            //BillToAddressLineFinal = BillToAddressLineFinal.PadLeft(BillToAddressLineFinal.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                            objStructF27.ContentsOfField = BillToAddressLineFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        clsModule.clsFinalValuesBillToAddress1 = BillToAddressLineFinal;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C0888A": // BillTo City Final
                        string BillToCityFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {
                            //if (clsModule.isMisscellenaousBillToCodePickup == "Y")
                            //{
                            string BillToDetailsMasterDBKeyCity = clsModule.OCR_Matched_MasterDB_KeyBillTo;
                            string BillToDetailsSearchEngineCity = clsModule.OCR_matched_SearchenagineBillTo;
                            string BillToDetailsMasterCity = clsModule.clsMasterValuesAPR001BillToDetails;
                            if (!string.IsNullOrEmpty(BillToDetailsMasterDBKeyCity) && BillToDetailsMasterDBKeyCity == "YES")
                            {
                                if (clsModule.isMailToAttachedShipper == "Y" && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail))
                                {
                                    string[] arrshipperDetilsMasterCell = clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail.Split('~');
                                    if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[2]))
                                    {
                                        BillToCityFinal = arrshipperDetilsMasterCell[2].ToString();
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                                else if (clsModule.isMailToAttachedConsignee == "Y" && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail))
                                {
                                    string[] arrshipperDetilsMasterCell = clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail.Split('~');
                                    if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[2]))
                                    {
                                        BillToCityFinal = arrshipperDetilsMasterCell[2].ToString();
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                                else if (clsModule.isMailToAttachedShipper != "Y" && clsModule.isMailToAttachedConsignee != "Y")
                                {
                                    #region BillToMailTo Code
                                    if (clsModule.isBillToAttachedShipper == "Y")
                                    {
                                        string[] arrshipperDetilsMasterCell = clsModule.ShipperPrepaidCodeDetail.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[2]))
                                        {
                                            BillToCityFinal = arrshipperDetilsMasterCell[2].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                    else if (clsModule.isBillToAttachedConsignee == "Y")
                                    {
                                        string[] arrshipperDetilsMasterCell = clsModule.ConsigneeCollectCodeDetail.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[2]))
                                        {
                                            BillToCityFinal = arrshipperDetilsMasterCell[2].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                    // If Bill to code matches with Shipper and consignee then drop the OCR Bill To OCR
                                    #endregion
                                }
                                else
                                {
                                    BillToCityFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToCity; ;
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "1F";
                                }
                            }
                            else if (!string.IsNullOrEmpty(BillToDetailsSearchEngineCity) && BillToDetailsSearchEngineCity == "YES")
                            {
                                BillToCityFinal = ClsSearchEngin.ClsSearchEnginBillToCity;
                                objStructF27.Flag = "1";
                                objStructF27.Remarks = "6F";
                            }
                            else
                            {
                                // TO DO => OCR
                                BillToCityFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToCity;
                                objStructF27.Flag = "0";
                                objStructF27.Remarks = "1F";
                            }
                            //}
                            //else
                            //{
                            //    // TO DO => OCR
                            //    BillToCityFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToCity; ;//Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(6, GridDataShipperCity));
                            //}
                        }

                        if (!string.IsNullOrEmpty(BillToCityFinal))
                        {
                            // BillToCityFinal = BillToCityFinal.PadLeft(BillToCityFinal.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                            objStructF27.ContentsOfField = BillToCityFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        clsModule.clsFinalValuesBillToCity = BillToCityFinal;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C0888B": // BillTo State Final
                        string BillToStateFinal = string.Empty;
                        //string DBBillToStateFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {
                            //if (clsModule.isMisscellenaousBillToCodePickup == "Y")
                            //{
                            string BillToDetailsMasterDBKeyState = clsModule.OCR_Matched_MasterDB_KeyBillTo;
                            string BillToDetailsSearchEngineSate = clsModule.OCR_matched_SearchenagineBillTo;
                            string BillToDetailsMasterState = clsModule.clsMasterValuesAPR001BillToDetails;
                            if (!string.IsNullOrEmpty(BillToDetailsMasterDBKeyState) && BillToDetailsMasterDBKeyState == "YES")
                            {
                                if (clsModule.isMailToAttachedShipper == "Y" && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail))
                                {
                                    string[] arrshipperDetilsMasterCell = clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail.Split('~');
                                    if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[3]))
                                    {
                                        BillToStateFinal = arrshipperDetilsMasterCell[3].ToString();
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                                else if (clsModule.isMailToAttachedConsignee == "Y" && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail))
                                {
                                    string[] arrshipperDetilsMasterCell = clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail.Split('~');
                                    if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[3]))
                                    {
                                        BillToStateFinal = arrshipperDetilsMasterCell[3].ToString();
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                                else if (clsModule.isMailToAttachedShipper != "Y" && clsModule.isMailToAttachedConsignee != "Y")
                                {
                                    #region BillToMailTo Code
                                    if (clsModule.isBillToAttachedShipper == "Y")
                                    {
                                        string[] arrshipperDetilsMasterCell = clsModule.ShipperPrepaidCodeDetail.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[3]))
                                        {
                                            BillToStateFinal = arrshipperDetilsMasterCell[3].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                    else if (clsModule.isBillToAttachedConsignee == "Y")
                                    {
                                        string[] arrshipperDetilsMasterCell = clsModule.ConsigneeCollectCodeDetail.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[3]))
                                        {
                                            BillToStateFinal = arrshipperDetilsMasterCell[3].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    BillToStateFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToState;
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "1F";
                                }
                            }
                            else if (!string.IsNullOrEmpty(BillToDetailsSearchEngineSate) && BillToDetailsSearchEngineSate == "YES")
                            {
                                BillToStateFinal = ClsSearchEngin.ClsSearchEnginBillToState;
                                objStructF27.Flag = "1";
                                objStructF27.Remarks = "6F";
                            }
                            else
                            {
                                // TO DO => OCR
                                BillToStateFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToState;
                                objStructF27.Flag = "0";
                                objStructF27.Remarks = "1F";
                            }
                            //}
                            //else
                            //{
                            //    // TO DO => OCR
                            //    BillToStateFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToState; ;//Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(7, GridDataShipperState));
                            //}
                        }

                        if (!string.IsNullOrEmpty(BillToStateFinal))
                        {
                            //BillToStateFinal = BillToStateFinal.PadLeft(BillToStateFinal.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                            objStructF27.ContentsOfField = BillToStateFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        clsModule.clsFinalValuesBillToState = BillToStateFinal;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C0888C": // BillTo Zip Final
                        string BillToZipFinal = string.Empty;
                        //string DBBillToZipFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {
                            //if (clsModule.isMisscellenaousBillToCodePickup == "Y")
                            //{
                            string BillToDetailsMasterDBKeyZip = clsModule.OCR_Matched_MasterDB_KeyBillTo;
                            string BillToDetailsSearchEngineZip = clsModule.OCR_matched_SearchenagineBillTo;
                            string BillToDetailsMasterZip = clsModule.clsMasterValuesAPR001BillToDetails;
                            if (!string.IsNullOrEmpty(BillToDetailsMasterDBKeyZip) && BillToDetailsMasterDBKeyZip == "YES")
                            {
                                if (clsModule.isMailToAttachedShipper == "Y" && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail))
                                {
                                    string[] arrshipperDetilsMasterCell = clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail.Split('~');
                                    if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[4]))
                                    {
                                        BillToZipFinal = arrshipperDetilsMasterCell[4].ToString();
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                                else if (clsModule.isMailToAttachedConsignee == "Y" && !string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail))
                                {
                                    string[] arrshipperDetilsMasterCell = clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail.Split('~');
                                    if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[4]))
                                    {
                                        BillToZipFinal = arrshipperDetilsMasterCell[4].ToString();
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                                else if (clsModule.isMailToAttachedShipper != "Y" && clsModule.isMailToAttachedConsignee != "Y")
                                {
                                    #region BillToMailTo Code
                                    if (clsModule.isBillToAttachedShipper == "Y")
                                    {
                                        string[] arrshipperDetilsMasterCell = clsModule.ShipperPrepaidCodeDetail.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[4]))
                                        {
                                            BillToZipFinal = arrshipperDetilsMasterCell[4].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                    else if (clsModule.isBillToAttachedConsignee == "Y")
                                    {
                                        string[] arrshipperDetilsMasterCell = clsModule.ConsigneeCollectCodeDetail.Split('~');
                                        if (!string.IsNullOrEmpty(arrshipperDetilsMasterCell[4]))
                                        {
                                            BillToZipFinal = arrshipperDetilsMasterCell[4].ToString();
                                            objStructF27.Flag = "1";
                                            objStructF27.Remarks = "5F";
                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    BillToZipFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToZip;
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "1F";
                                }
                            }
                            else if (!string.IsNullOrEmpty(BillToDetailsSearchEngineZip) && BillToDetailsSearchEngineZip == "YES")
                            {
                                BillToZipFinal = ClsSearchEngin.ClsSearchEnginBillToZip;
                                objStructF27.Flag = "1";
                                objStructF27.Remarks = "6F";
                            }
                            else
                            {
                                // TO DO => OCR
                                BillToZipFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToZip;
                                objStructF27.Flag = "0";
                                objStructF27.Remarks = "1F";
                            }
                            //}
                            //else
                            //{
                            //    // TO DO => OCR
                            //    BillToZipFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToZip; //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(8, GridDataShipperZip));
                            //}
                        }
                        if (!string.IsNullOrEmpty(BillToZipFinal))
                        {
                            //BillToZipFinal = BillToZipFinal.PadLeft(BillToZipFinal.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                            objStructF27.ContentsOfField = BillToZipFinal;
                        }
                        else
                        {
                            objStructF27.ContentsOfField = "*****";
                        }
                        clsModule.clsFinalValuesBillToZip = BillToZipFinal;
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                    case "C0888D": // BillTo Telephone Number Final
                        string BillToTelFinal = string.Empty; //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(9, GridDataShipperTel));
                        //string DBBillToTelFinal = string.Empty;
                        if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1) && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "*****" && clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim() != "@@@@@")
                        {
                            //if (clsModule.isMisscellenaousBillToCodePickup == "Y")
                            //{
                            string BillToDetailsMasterDBKeyTelePhoneNo = clsModule.OCR_Matched_MasterDB_KeyBillTo;
                            string BillToDetailsSearchEngineTelePhoneNo = clsModule.OCR_matched_SearchenagineBillTo;
                            // Master Engine
                            if (!string.IsNullOrEmpty(BillToDetailsMasterDBKeyTelePhoneNo) && BillToDetailsMasterDBKeyTelePhoneNo == "YES")
                            {
                                if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001BillToTelePhoneNo) && !string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo))
                                {
                                    objRet58 = ObjSbr.IQSBR058(Module1.func_RemoveSpecialCharacter(clsModule.clsMasterValuesAPR001BillToTelePhoneNo), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo), true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
                                    if (objRet58.CharChanged <= 1)
                                    {
                                        BillToTelFinal = clsModule.clsMasterValuesAPR001BillToTelePhoneNo;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "5F";
                                    }
                                }
                                else
                                {
                                    if (clsF3.ClsOCRValues.BillToTelephoneOCRWithHighConfidence == "Y")
                                    {
                                        BillToTelFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "1F";
                                    }
                                    else
                                    {
                                        BillToTelFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }
                                }

                            }
                            // Search Engine 
                            else if (!string.IsNullOrEmpty(BillToDetailsSearchEngineTelePhoneNo) && BillToDetailsSearchEngineTelePhoneNo == "YES")
                            {
                                objRet58 = ObjSbr.IQSBR058(Module1.func_RemoveSpecialCharacter(ClsSearchEngin.ClsSearchEnginBillToTelePhoneNo), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo), true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
                                if (objRet58.CharChanged <= 1)
                                {
                                    BillToTelFinal = ClsSearchEngin.ClsSearchEnginBillToTelePhoneNo;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "6F";
                                }
                                else
                                {
                                    if (clsF3.ClsOCRValues.BillToTelephoneOCRWithHighConfidence == "Y")
                                    {
                                        BillToTelFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo;
                                        objStructF27.Flag = "1";
                                        objStructF27.Remarks = "1F";
                                    }
                                    else
                                    {
                                        BillToTelFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo;
                                        objStructF27.Flag = "0";
                                        objStructF27.Remarks = "1F";
                                    }

                                }
                            }
                            // OCR
                            else
                            {
                                if (clsF3.ClsOCRValues.BillToTelephoneOCRWithHighConfidence == "Y")
                                {
                                    BillToTelFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo;
                                    objStructF27.Flag = "1";
                                    objStructF27.Remarks = "1F";
                                }
                                else
                                {
                                    BillToTelFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo;
                                    objStructF27.Flag = "0";
                                    objStructF27.Remarks = "1F";
                                }
                            }
                            //}
                            //else
                            //{
                            //    // TO DO => OCR
                            //    BillToTelFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo; //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(9, GridDataShipperTel));
                            //}
                        }

                        if (!string.IsNullOrEmpty(BillToTelFinal) && BillToTelFinal != "*****" && BillToTelFinal != "@@@@@" && BillToTelFinal != "0")
                        {
                            //BillToTelFinal = BillToTelFinal.PadLeft(BillToTelFinal.Length, '0'); //"9".PadLeft(ShipperDetails.Length, '9');
                            objStructF27.ContentsOfField = BillToTelFinal;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo))
                            {
                                BillToTelFinal = clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo;
                                objStructF27.ContentsOfField = BillToTelFinal;
                                objStructF27.Flag = "0";
                                objStructF27.Remarks = "1F";
                            }
                            else
                            {
                                objStructF27.ContentsOfField = "*****";
                            }
                        }
                        objStructF27.X1 = 10;
                        objStructF27.Y1 = 10;
                        objStructF27.X2 = 20;
                        objStructF27.Y2 = 20;
                        objStructF27.PageNo = 998;
                        objStructF27.Status = "S";
                        break;

                }
            }
            catch (Exception)
            {
                //MessageBox.Show("F27 - " + ex.Message.ToString());
            }
            return objStructF27;
        }

        public void GetShipperDetailsMaster(string shipperCodePickup, string strflag, string strRAT, ref string strShipperCodeMasterOrTelNoOCR, ref string strShipperDetailsMaster)
        {
            if (strflag == "S")
            {
                clsModule.isCMBTPCMBTCpresentShipper = string.Empty;
                clsModule.isCMBTPCMBTCpresentValueShipper = string.Empty;
                clsModule.ShipperPrepaidCodeDetail = string.Empty;
                clsModule.ShipperCollectCodeDetail = string.Empty; // NOT IN USE
                // Shipper 
                clsModule.clsMasterValuesAPR001ShipperRATCode = string.Empty;
                clsModule.clsMasterValuesAPR001ShipperRATDetails = string.Empty;
                clsModule.clsMasterValuesAPR001ShipperRATCodeFinal = string.Empty;
                clsModule.clsMasterValuesAPR001ShipperRATTelePhoneNo = string.Empty;

                // BS
                clsModule.clsPickupValuesFRP001BSMatched = string.Empty;
                clsModule.clsPickupValuesFRP001BSMatchedCode = string.Empty;
                clsModule.clsPickupValuesFRP001BSMatchedDetails = string.Empty;

                // Mail To attached
                clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail = string.Empty;

                // Address1 and address2 with pipe
                clsModule.clsMasterValuesAPR001ShipperDetailsWithPipe = string.Empty;
                //clsModule.clsMasterValuesAPR001ShipperRATDetailsWithPipe = string.Empty;
                //clsModule.clsMasterValuesAPR001ShipperBSDetailsWithPipe = string.Empty;
                clsModule.clsMasterValuesAPR001MailToAttachedShipperWithPipe = string.Empty;
                clsModule.clsMasterValuesAPR001BillToAttachedShipperWithPipe = string.Empty;

            }
            else if (strflag == "C")
            {
                clsModule.isCMBTPCMBTCpresent = "N";
                clsModule.isCMBTPCMBTCpresentValue = string.Empty;
                clsModule.ConsigneePrepaidCodeDetail = string.Empty; // NOT IN USE
                clsModule.ConsigneeCollectCodeDetail = string.Empty;
                // consignee
                clsModule.clsMasterValuesAPR001ConsigneeHATCode = string.Empty;
                clsModule.clsMasterValuesAPR001ConsigneeHATDetails = string.Empty;
                clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal = string.Empty;
                clsModule.clsMasterValuesAPR001ConsigneeHATTelePhoneNo = string.Empty;

                // Mail To attached
                clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail = string.Empty;

                // Address1 and address2 with pipe
                clsModule.clsMasterValuesAPR001ConsigneeDetailsWithPipe = string.Empty;
                //clsModule.clsMasterValuesAPR001ConsigneeHATDetailsWithPipe = string.Empty;
                clsModule.clsMasterValuesAPR001MailToAttachedConsigneeWithPipe = string.Empty;
                clsModule.clsMasterValuesAPR001BillToAttachedConsigneeWithPipe = string.Empty;

            }

            //RetStructIQSBR026 objRet26 = new RetStructIQSBR026();
            RetStructIQSBR050 objRet50 = new RetStructIQSBR050();
            ClsCNC objCNC = new ClsCNC();
            clsCNCSBR ObjSbr = new clsCNCSBR();
            string shipperTelNoPickup = String.Empty;
            bool isshippermisscellaneous = false;
            bool isconsigneemisscellaneous = false;
            bool isBillTomisscellaneous = false;
            // Check CMCLS flag => customer class condition 
            string customerClassCondition = "";
            if (strflag == "S")
            {
                customerClassCondition = " AND CMCLS = 'F'";
            }
            else if (strflag == "C")
            {
                customerClassCondition = " AND CMCLS = 'F'";
            }
            else
            {
                customerClassCondition = " AND (CMCLS = 'F' OR CMCLS = 'B')";
            }
            // misscellaneous then check with the telephone number
            if (strflag == "S")
            {
                if (shipperCodePickup.StartsWith("99") || shipperCodePickup.StartsWith("98") || shipperCodePickup.StartsWith("97"))
                {
                    isshippermisscellaneous = true;
                    shipperTelNoPickup = clsF3.ClsOCRValues.ClsOCRValuesShipperTelePhoneNo;
                }
            }
            else if (strflag == "C")
            {
                if (shipperCodePickup.StartsWith("98") || shipperCodePickup.StartsWith("99") || shipperCodePickup.StartsWith("97"))
                {
                    isconsigneemisscellaneous = true;
                    shipperTelNoPickup = clsF3.ClsOCRValues.ClsOCRValuesConsigneeTelePhoneNo;
                }
            }
            else
            {
                if (shipperCodePickup.StartsWith("97") || shipperCodePickup.StartsWith("99") || shipperCodePickup.StartsWith("98"))
                {
                    isBillTomisscellaneous = true;
                    shipperTelNoPickup = clsF3.ClsOCRValues.ClsOCRValuesBillToTelePhoneNo;
                }
            }
            // Check shipper => RAT
            if (strflag == "S" && strRAT == "Y")
            {
                string strName = string.Empty;
                if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesShipperName) && clsF3.ClsOCRValues.ClsOCRValuesShipperName != "*****")
                {
                    string[] arrName = clsF3.ClsOCRValues.ClsOCRValuesShipperName.Split(' ');
                    if (arrName.Length > 0)
                    {
                        if (arrName[0].ToString().Length > 2)
                        {
                            strName = arrName[0];
                        }
                        else
                        {
                            if (arrName.Length >= 2)
                            {
                                strName = arrName[0] + " " + arrName[1];
                            }
                        }

                    }

                    //string queryStringMaster = "select CMCUST,CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP,CMPHON from ARP001 where  CMADR1 like '%EXIT SAIA%' and CMNAME like'%" + strName + "%' and CMTID= '"+ clsModule.clsPickupValuesFRP001ShipperTerminalID + "' ";
                    string queryStringMaster = "select CMCUST,CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP,CMPHON from ARP001 where CMSTAT = 'A' AND CMADR1 like '%EXIT SAIA%' and CMNAME like'" + strName.ToUpper() + "%' and CMTID= '" + clsModule.clsPickupValuesFRP001ShipperTerminalID + "'" + customerClassCondition + "";
                    DataTable dtMaster_Shipper_Details = GetDataTableByText(As400_ConnectionString, queryStringMaster);
                    //DataTable dtMaster_Shipper_Details = GetDataTableByText(queryStringMaster);
                    if (dtMaster_Shipper_Details != null)
                    {
                        if (dtMaster_Shipper_Details.Rows.Count == 1)
                        {
                            string CMNAME = string.Empty;
                            CMNAME = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim();
                            objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(CMNAME), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperName.Trim()));
                            if (objRet50.PercentageMatch > 80)
                            {
                                strShipperCodeMasterOrTelNoOCR = "RAT - " + dtMaster_Shipper_Details.Rows[0]["CMCUST"].ToString().Trim();
                                if (dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim())
                                {
                                    strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                                }
                                else
                                {
                                    strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                                }
                                clsModule.clsMasterValuesAPR001ShipperRATCode = strShipperCodeMasterOrTelNoOCR;
                                clsModule.clsMasterValuesAPR001ShipperRATDetails = strShipperDetailsMaster;
                                clsModule.clsMasterValuesAPR001ShipperRATCodeFinal = dtMaster_Shipper_Details.Rows[0]["CMCUST"].ToString().Trim(); //strShipperCodeMasterOrTelNoOCR;
                                clsModule.clsMasterValuesAPR001ShipperRATTelePhoneNo = dtMaster_Shipper_Details.Rows[0]["CMPHON"].ToString().Trim();
                                shipperCodePickup = dtMaster_Shipper_Details.Rows[0]["CMCUST"].ToString().Trim();
                                //clsModule.clsMasterValuesAPR001ShipperRATDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                                clsModule.clsMasterValuesAPR001ShipperDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                            }
                        }
                        else
                        {
                            if (dtMaster_Shipper_Details.Rows.Count > 0)
                            {
                                bool isRecordFound = false;
                                for (int i = 0; i < dtMaster_Shipper_Details.Rows.Count; i++)
                                {
                                    string CMNAME = string.Empty;
                                    if (!isRecordFound)
                                    {
                                        if (!DBNull.Value.Equals(dtMaster_Shipper_Details.Rows[i]["CMNAME"]))
                                        {
                                            CMNAME = dtMaster_Shipper_Details.Rows[i]["CMNAME"].ToString().Trim();
                                            objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(CMNAME), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesShipperName.Trim()));
                                            if (objRet50.PercentageMatch > 80)
                                            {
                                                strShipperCodeMasterOrTelNoOCR = "RAT - " + dtMaster_Shipper_Details.Rows[i]["CMCUST"].ToString().Trim();
                                                if (dtMaster_Shipper_Details.Rows[i]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[i]["CMADR2"].ToString().Trim())
                                                {
                                                    strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[i]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMZIP"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[i]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[i]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMZIP"].ToString().Trim();
                                                }
                                                clsModule.clsMasterValuesAPR001ShipperRATCode = strShipperCodeMasterOrTelNoOCR;
                                                clsModule.clsMasterValuesAPR001ShipperRATDetails = strShipperDetailsMaster;
                                                clsModule.clsMasterValuesAPR001ShipperRATCodeFinal = dtMaster_Shipper_Details.Rows[i]["CMCUST"].ToString().Trim(); //strShipperCodeMasterOrTelNoOCR;
                                                shipperCodePickup = dtMaster_Shipper_Details.Rows[i]["CMCUST"].ToString().Trim();
                                                //clsModule.clsMasterValuesAPR001ShipperRATDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                                                clsModule.clsMasterValuesAPR001ShipperDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                                                if (dtMaster_Shipper_Details.Rows[i]["CMPHON"].ToString().Trim() != "0")
                                                {
                                                    clsModule.clsMasterValuesAPR001ShipperRATTelePhoneNo = dtMaster_Shipper_Details.Rows[i]["CMPHON"].ToString().Trim();
                                                }

                                                isRecordFound = true;
                                                break;
                                            }

                                        }
                                    }
                                }
                            }
                            else
                            {
                                strShipperCodeMasterOrTelNoOCR = "RAT";
                                strShipperDetailsMaster = "";
                                clsModule.clsMasterValuesAPR001ShipperRATCode = strShipperCodeMasterOrTelNoOCR;
                                clsModule.clsMasterValuesAPR001ShipperRATDetails = strShipperDetailsMaster;
                                clsModule.clsMasterValuesAPR001ShipperRATTelePhoneNo = "";
                                //clsModule.clsMasterValuesAPR001ShipperRATDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                                if (dtMaster_Shipper_Details.Rows.Count > 0)
                                {
                                    clsModule.clsMasterValuesAPR001ShipperDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                                }
                            }
                        }
                    }
                    if (dtMaster_Shipper_Details != null)
                    {
                        dtMaster_Shipper_Details.Dispose();
                        dtMaster_Shipper_Details = null;
                    }
                }
            }
            // Check Consignee => HAT
            else if (strflag == "C" && strRAT == "Y")
            {
                string strName = string.Empty;
                if (!string.IsNullOrEmpty(clsF3.ClsOCRValues.ClsOCRValuesConsigneeName) && clsF3.ClsOCRValues.ClsOCRValuesConsigneeName != "*****")
                {
                    string[] arrName = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName.Split(' ');
                    if (arrName.Length > 0)
                    {
                        if (arrName[0].ToString().Length > 2)
                        {
                            strName = arrName[0];
                        }
                        else
                        {
                            if (arrName.Length >= 2)
                            {
                                strName = arrName[0] + " " + arrName[1];
                            }
                        }

                    }


                    //string queryStringMaster = "select CMCUST,CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP,CMPHON from ARP001 where  CMADR1 like '%EXIT SAIA%' and CMNAME like'%" + strName + "%' and CMTID= '"+ clsModule.clsPickupValuesFRP001ShipperTerminalID + "' ";
                    string queryStringMaster = "select CMCUST,CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP,CMPHON from ARP001 where CMSTAT = 'A' AND CMADR1 like '%% SAIA%' and CMNAME like'" + strName.ToUpper() + "%' and CMTID= '" + clsModule.clsPickupValuesFRP001ConsigneeTerminalID + "'" + customerClassCondition + "";

                    // DataTable dtMaster_Shipper_Details = GetDataTableByText(queryStringMaster);
                    DataTable dtMaster_Shipper_Details = GetDataTableByText(As400_ConnectionString, queryStringMaster);
                    if (dtMaster_Shipper_Details != null)
                    {
                        if (dtMaster_Shipper_Details.Rows.Count == 1)
                        {
                            string CMNAME = string.Empty;
                            CMNAME = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim();
                            objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(CMNAME), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesConsigneeName.Trim()));
                            if (objRet50.PercentageMatch > 80)
                            {

                                strShipperCodeMasterOrTelNoOCR = "HAT - " + dtMaster_Shipper_Details.Rows[0]["CMCUST"].ToString().Trim();
                                if (dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim())
                                {
                                    strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                                }
                                else
                                {
                                    strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                                }

                                clsModule.clsMasterValuesAPR001ConsigneeHATCode = strShipperCodeMasterOrTelNoOCR;
                                clsModule.clsMasterValuesAPR001ConsigneeHATDetails = strShipperDetailsMaster;
                                clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal = dtMaster_Shipper_Details.Rows[0]["CMCUST"].ToString().Trim(); //strShipperCodeMasterOrTelNoOCR;
                                shipperCodePickup = dtMaster_Shipper_Details.Rows[0]["CMCUST"].ToString().Trim();
                                clsModule.clsMasterValuesAPR001ConsigneeHATTelePhoneNo = dtMaster_Shipper_Details.Rows[0]["CMPHON"].ToString().Trim();
                                //clsModule.clsMasterValuesAPR001ConsigneeHATDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                                clsModule.clsMasterValuesAPR001ConsigneeDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                            }
                        }
                        else
                        {
                            if (dtMaster_Shipper_Details.Rows.Count > 0)
                            {
                                bool isRecordFound = false;
                                for (int i = 0; i < dtMaster_Shipper_Details.Rows.Count; i++)
                                {
                                    string CMNAME = string.Empty;
                                    if (!isRecordFound)
                                    {
                                        if (!DBNull.Value.Equals(dtMaster_Shipper_Details.Rows[i]["CMNAME"]))
                                        {
                                            CMNAME = dtMaster_Shipper_Details.Rows[i]["CMNAME"].ToString().Trim();
                                            objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(CMNAME), Module1.func_RemoveSpecialCharacter(clsF3.ClsOCRValues.ClsOCRValuesConsigneeName.Trim()));
                                            if (objRet50.PercentageMatch > 80)
                                            {
                                                strShipperCodeMasterOrTelNoOCR = "HAT - " + dtMaster_Shipper_Details.Rows[i]["CMCUST"].ToString().Trim();
                                                if (dtMaster_Shipper_Details.Rows[i]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[i]["CMADR2"].ToString().Trim())
                                                {
                                                    strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[i]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMZIP"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[i]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[i]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[i]["CMZIP"].ToString().Trim();
                                                }
                                                clsModule.clsMasterValuesAPR001ConsigneeHATCode = strShipperCodeMasterOrTelNoOCR;
                                                clsModule.clsMasterValuesAPR001ConsigneeHATDetails = strShipperDetailsMaster;
                                                clsModule.clsMasterValuesAPR001ConsigneeHATCodeFinal = dtMaster_Shipper_Details.Rows[i]["CMCUST"].ToString().Trim(); //strShipperCodeMasterOrTelNoOCR;
                                                shipperCodePickup = dtMaster_Shipper_Details.Rows[i]["CMCUST"].ToString().Trim();
                                                //clsModule.clsMasterValuesAPR001ConsigneeHATDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                                                clsModule.clsMasterValuesAPR001ConsigneeDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                                                if (dtMaster_Shipper_Details.Rows[i]["CMPHON"].ToString().Trim() != "0")
                                                {
                                                    clsModule.clsMasterValuesAPR001ConsigneeHATTelePhoneNo = dtMaster_Shipper_Details.Rows[i]["CMPHON"].ToString().Trim();
                                                }

                                                isRecordFound = true;
                                                break;
                                            }

                                        }
                                    }
                                }
                            }
                            else
                            {
                                strShipperCodeMasterOrTelNoOCR = "HAT";
                                strShipperDetailsMaster = "";
                                clsModule.clsMasterValuesAPR001ConsigneeHATCode = strShipperCodeMasterOrTelNoOCR;
                                clsModule.clsMasterValuesAPR001ConsigneeHATDetails = strShipperDetailsMaster;
                                clsModule.clsMasterValuesAPR001ConsigneeHATTelePhoneNo = "";
                                //clsModule.clsMasterValuesAPR001ConsigneeHATDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                                clsModule.clsMasterValuesAPR001ConsigneeDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                            }
                        }
                    }
                    if (dtMaster_Shipper_Details != null)
                    {
                        dtMaster_Shipper_Details.Dispose();
                        dtMaster_Shipper_Details = null;
                    }
                }
            }

            // Normal case for Shipper, Consignee and Bill-To
            else
            {
                if (isshippermisscellaneous == true || isconsigneemisscellaneous == true || isBillTomisscellaneous == true)
                {
                    #region NOT IN USE => No Need to use telephone for search // as discussed code enabled on 5-Feb-2021
                    shipperTelNoPickup = Module1.func_RemoveSpecialCharacter(shipperTelNoPickup);
                    if (!string.IsNullOrEmpty(shipperTelNoPickup) && shipperTelNoPickup != "*****")
                    {
                        string queryStringMaster = "select CMCUST,CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP from ARP001 where CMSTAT = 'A' AND CMPHON='" + shipperTelNoPickup.Trim().Replace("X", "") + "'" + customerClassCondition + "";
                        DataTable dtMaster_Shipper_Details = GetDataTableByText(As400_ConnectionString, queryStringMaster);

                        if (dtMaster_Shipper_Details.Rows.Count > 0)
                        {
                            strShipperCodeMasterOrTelNoOCR = shipperTelNoPickup;
                            if (dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim())
                            {
                                strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                            }
                            else
                            {
                                strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                            }
                            if (strflag == "S")
                            {
                                clsModule.clsMasterValuesAPR001ShipperDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                                clsModule.clsMasterValuesAPR001ShipperCodeUsingTelphone = dtMaster_Shipper_Details.Rows[0]["CMCUST"].ToString().Trim();
                            }
                            else if (strflag == "C")
                            {
                                clsModule.clsMasterValuesAPR001ConsigneeDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                                clsModule.clsMasterValuesAPR001ConsigneeCodeUsingTelphone = dtMaster_Shipper_Details.Rows[0]["CMCUST"].ToString().Trim();
                            }
                            else
                            {
                                clsModule.clsMasterValuesAPR001BillToCodeUsingTelphone = dtMaster_Shipper_Details.Rows[0]["CMCUST"].ToString().Trim();
                            }

                        }
                        if (dtMaster_Shipper_Details != null)
                        {
                            dtMaster_Shipper_Details.Dispose();
                            dtMaster_Shipper_Details = null;
                        }
                    }
                    #endregion
                }
                // NON miscellaneous
                else
                {
                    //string queryStringMaster = "select CMCUST,CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP,CMBTP,CMBTC from ARP001 where CMCUST='" + shipperCodePickup + "'";
                    // Mailing address are added
                    string queryStringMaster = "select CMCUST,CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP,CMBTP,CMBTC,CMMAD1,CMMAD2,CMMCIT,CMMST,CMMZIP from ARP001 where CMSTAT = 'A' AND CMCUST='" + shipperCodePickup + "'" + customerClassCondition + "";
                    DataTable dtMaster_Shipper_Details = GetDataTableByText(As400_ConnectionString, queryStringMaster);

                    if (dtMaster_Shipper_Details != null && dtMaster_Shipper_Details.Rows.Count > 0)
                    {
                        strShipperCodeMasterOrTelNoOCR = dtMaster_Shipper_Details.Rows[0]["CMCUST"].ToString().Trim();
                        if (dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim())
                        {
                            strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                        }
                        else
                        {
                            strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                        }
                        // Address1 and address2 with pipe
                        if (strflag == "S")
                        {
                            clsModule.clsMasterValuesAPR001ShipperDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                        }
                        else if (strflag == "C")
                        {
                            clsModule.clsMasterValuesAPR001ConsigneeDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                        }


                        string CMBTP = string.Empty;
                        string CMBTC = string.Empty;
                        CMBTP = dtMaster_Shipper_Details.Rows[0]["CMBTP"].ToString().Trim();
                        CMBTC = dtMaster_Shipper_Details.Rows[0]["CMBTC"].ToString().Trim();
                        // Get the CMBTP and CMBTC Details
                        #region "Code commented on 4-March-2021 => MailToAttach and BillToAttach now we are fetching this details at BillToAttach Code routine"
                        if (strflag == "C")
                        {
                            // Get Mailing address attached to Consignee
                            //if (!string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMAD1"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMAD2"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMCIT"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMST"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMZIP"].ToString().Trim()))
                            if (!string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMCIT"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMST"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMZIP"].ToString().Trim()))
                            {
                                clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail = dtMaster_Shipper_Details.Rows[0]["CMMAD1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMAD2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMCIT"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMZIP"].ToString().Trim();
                            }
                            // address1 and address2 with pipe
                            clsModule.clsMasterValuesAPR001MailToAttachedConsigneeWithPipe = dtMaster_Shipper_Details.Rows[0]["CMMAD1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMMAD2"].ToString().Trim();

                            //if (string.IsNullOrEmpty(CMBTP) && string.IsNullOrEmpty(CMBTC))
                            if (string.IsNullOrEmpty(CMBTC))
                            {
                                clsModule.isCMBTPCMBTCpresent = "N";
                                //clsModule.isCMBTPCMBTCpresentValue = CMBTP + "~" + CMBTC;
                                clsModule.isCMBTPCMBTCpresentValue = CMBTC;
                            }
                            //if (!string.IsNullOrEmpty(CMBTP) || !string.IsNullOrEmpty(CMBTC))
                            if (!string.IsNullOrEmpty(CMBTC))
                            {
                                clsModule.isCMBTPCMBTCpresent = "Y";
                                //clsModule.isCMBTPCMBTCpresentValue = CMBTP + "~" + CMBTC;
                                clsModule.isCMBTPCMBTCpresentValue = CMBTC;
                            }

                        }
                        #endregion
                        else if (strflag == "S") //if (strflag == "S")
                        {
                            // Get Mailing address attached to shipper
                            //if (!string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMAD1"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMAD2"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMCIT"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMST"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMZIP"].ToString().Trim()))
                            if (!string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMCIT"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMST"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMZIP"].ToString().Trim()))
                            {
                                clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail = dtMaster_Shipper_Details.Rows[0]["CMMAD1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMAD2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMCIT"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMZIP"].ToString().Trim();
                            }
                            // address1 and address2 with pipe
                            clsModule.clsMasterValuesAPR001MailToAttachedShipperWithPipe = dtMaster_Shipper_Details.Rows[0]["CMMAD1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMMAD2"].ToString().Trim();
                            //if (string.IsNullOrEmpty(CMBTP) && string.IsNullOrEmpty(CMBTC))
                            if (string.IsNullOrEmpty(CMBTP))
                            {
                                clsModule.isCMBTPCMBTCpresentShipper = "N";
                                //clsModule.isCMBTPCMBTCpresentValueShipper = CMBTP + "~" + CMBTC;
                                clsModule.isCMBTPCMBTCpresentValueShipper = CMBTP;
                            }
                            //if (!string.IsNullOrEmpty(CMBTP) || !string.IsNullOrEmpty(CMBTC))
                            if (!string.IsNullOrEmpty(CMBTP))
                            {
                                clsModule.isCMBTPCMBTCpresentShipper = "Y";
                                //clsModule.isCMBTPCMBTCpresentValueShipper = CMBTP + "~" + CMBTC;
                                clsModule.isCMBTPCMBTCpresentValueShipper = CMBTP;
                            }
                            // STEP:- 3 check for the BS
                            if (clsModule.clsPickupValuesFRP001BS == "Y")
                            {
                                // Check address 
                                string[] arrshipperDetilsMasterCell = strShipperDetailsMaster.Split('~');

                                string AddressLine1 = clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1;
                                string City = clsF3.ClsOCRValues.ClsOCRValuesShipperCity;
                                string State = clsF3.ClsOCRValues.ClsOCRValuesShipperState;
                                string Zip = clsF3.ClsOCRValues.ClsOCRValuesShipperZip;
                                string shipperDetilsMasterCellOCR = AddressLine1 + " " + City + " " + State + " " + Zip;
                                string shipperDetilsMasterCellMst = arrshipperDetilsMasterCell[1] + " " + arrshipperDetilsMasterCell[2] + " " + arrshipperDetilsMasterCell[3] + " " + arrshipperDetilsMasterCell[4];
                                string IsMatchAddress = MatchAddress(shipperDetilsMasterCellOCR, shipperDetilsMasterCellMst, "S");

                                //// Check for the Blind Shipper 
                                ////if (isAddressLine1 && isCity && isZip) Removed matching crietria city from the shiiper
                                //if (((isAddressLine1 && isCity && isState && isZip) || (isAddressLine1 && isZip && isState))) // ignore city TO DO zip also
                                if (IsMatchAddress == "Y")
                                {
                                    strShipperCodeMasterOrTelNoOCR = "BS - " + dtMaster_Shipper_Details.Rows[0]["CMCUST"].ToString().Trim();
                                    if (dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim())
                                    {
                                        strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                                    }
                                    else
                                    {
                                        strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                                    }

                                    clsModule.clsPickupValuesFRP001BSMatched = "Y";
                                }
                                else
                                {
                                    strShipperCodeMasterOrTelNoOCR = "BS - ";
                                    if (dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim())
                                    {
                                        strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                                    }
                                    else
                                    {
                                        strShipperDetailsMaster = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                                    }
                                    clsModule.clsPickupValuesFRP001BSMatched = "N";
                                }
                                // address1 and address2 with pipe
                                clsModule.clsMasterValuesAPR001ShipperDetailsWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();

                            }

                        }

                        // CMBTP == CMBTC => Then execute the same query for prepaid and collect  1) Get Shipper Prepaid Code Detail OR Consignee Prepaid Code Details  2) Get Shipper Collect Code Detail OR Consignee Collect Code Details
                        if (!string.IsNullOrEmpty(CMBTP.Trim()) && !string.IsNullOrEmpty(CMBTC.Trim()) && (CMBTP.Trim() == CMBTC.Trim()))
                        {
                            //if (CMBTP.Trim() == CMBTC.Trim())
                            //{
                            string queryStringShipperConsigneePrepaid = "select CMCUST,CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP,CMBTP,CMBTC from ARP001 where CMSTAT = 'A' AND CMCUST='" + shipperCodePickup + "' and CMBTP = '" + CMBTP + "'" + customerClassCondition + "";
                            DataTable dtMasterShipperConsigneePrepaidDetails = GetDataTableByText(As400_ConnectionString, queryStringShipperConsigneePrepaid);
                            if (dtMasterShipperConsigneePrepaidDetails.Rows.Count > 0)
                            {
                                if (strflag == "S")
                                {
                                    if (dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() == dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim())
                                    {
                                        clsModule.ShipperPrepaidCodeDetail = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMZIP"].ToString().Trim();
                                    }
                                    else
                                    {
                                        clsModule.ShipperPrepaidCodeDetail = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMZIP"].ToString().Trim();
                                    }

                                    //clsModule.ShipperCollectCodeDetail = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMZIP"].ToString().Trim();
                                    clsModule.clsMasterValuesAPR001BillToAttachedShipperWithPipe = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim();
                                }
                                else if (strflag == "C")
                                {
                                    if (dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() == dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim())
                                    {
                                        clsModule.ConsigneeCollectCodeDetail = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMZIP"].ToString().Trim();
                                    }
                                    else
                                    {
                                        clsModule.ConsigneeCollectCodeDetail = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMZIP"].ToString().Trim();
                                    }
                                    //clsModule.ConsigneePrepaidCodeDetail = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMZIP"].ToString().Trim();
                                    clsModule.clsMasterValuesAPR001BillToAttachedConsigneeWithPipe = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim();
                                }
                            }
                            //}
                        }
                        // CMBTP != CMBTC => Then execute the separate query for prepaid and collect
                        else
                        {
                            // Get the Shipper Prepaid Code Detail OR Consignee Prepaid Code Details
                            if (!string.IsNullOrEmpty(CMBTP))
                            {
                                string queryStringShipperConsigneePrepaid = "select CMCUST,CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP,CMBTP,CMBTC from ARP001 where CMSTAT = 'A' AND CMCUST='" + shipperCodePickup + "' and CMBTP = '" + CMBTP + "'" + customerClassCondition + "";
                                DataTable dtMasterShipperConsigneePrepaidDetails = GetDataTableByText(As400_ConnectionString, queryStringShipperConsigneePrepaid);
                                if (dtMasterShipperConsigneePrepaidDetails.Rows.Count > 0)
                                {
                                    if (strflag == "S")
                                    {
                                        if (dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() == dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim())
                                        {
                                            clsModule.ShipperPrepaidCodeDetail = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMZIP"].ToString().Trim();
                                        }
                                        else
                                        {
                                            clsModule.ShipperPrepaidCodeDetail = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMZIP"].ToString().Trim();
                                        }
                                        clsModule.clsMasterValuesAPR001BillToAttachedShipperWithPipe = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim();

                                    }
                                    //else if (strflag == "C")
                                    //{
                                    //    clsModule.ConsigneePrepaidCodeDetail = dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneePrepaidDetails.Rows[0]["CMZIP"].ToString().Trim();

                                    //}
                                }
                            }
                            // Get the Shipper Collect Code Detail OR Consignee Collect Code Details 
                            if (!string.IsNullOrEmpty(CMBTC))
                            {
                                string queryStringShipperConsigneeCollect = "select CMCUST,CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP,CMBTP,CMBTC from ARP001 where CMSTAT = 'A' AND CMCUST='" + shipperCodePickup + "' and CMBTC = '" + CMBTC + "'" + customerClassCondition + "";
                                DataTable dtMasterShipperConsigneeCollectDetails = GetDataTableByText(As400_ConnectionString, queryStringShipperConsigneeCollect);
                                if (dtMasterShipperConsigneeCollectDetails.Rows.Count > 0)
                                {
                                    //if (strflag == "S")
                                    //{
                                    //    clsModule.ShipperCollectCodeDetail = dtMasterShipperConsigneeCollectDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMZIP"].ToString().Trim();

                                    //}
                                    //else 
                                    if (strflag == "C")
                                    {
                                        if (dtMasterShipperConsigneeCollectDetails.Rows[0]["CMADR1"].ToString().Trim() == dtMasterShipperConsigneeCollectDetails.Rows[0]["CMADR2"].ToString().Trim())
                                        {
                                            clsModule.ConsigneeCollectCodeDetail = dtMasterShipperConsigneeCollectDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMZIP"].ToString().Trim();
                                        }
                                        else
                                        {
                                            clsModule.ConsigneeCollectCodeDetail = dtMasterShipperConsigneeCollectDetails.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMST"].ToString().Trim() + "~" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMZIP"].ToString().Trim();
                                        }
                                        clsModule.clsMasterValuesAPR001BillToAttachedConsigneeWithPipe = dtMasterShipperConsigneeCollectDetails.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMasterShipperConsigneeCollectDetails.Rows[0]["CMADR2"].ToString().Trim();
                                    }
                                }
                            }
                        }
                    }
                    if (dtMaster_Shipper_Details != null)
                    {
                        dtMaster_Shipper_Details.Dispose();
                        dtMaster_Shipper_Details = null;
                    }
                }
            }
            // Master Data ARP001 => Assign data to the global properties 
            if (strflag == "S")
            {
                if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ShipperCodeUsingTelphone))
                {
                    clsModule.clsMasterValuesAPR001ShipperCode = clsModule.clsMasterValuesAPR001ShipperCodeUsingTelphone;
                }
                else
                {
                    clsModule.clsMasterValuesAPR001ShipperCode = shipperCodePickup; //strShipperCodeMasterOrTelNoOCR;//
                }
                clsModule.clsMasterValuesAPR001ShipperTelePhoneNo = shipperTelNoPickup;
                clsModule.clsMasterValuesAPR001ShipperDetails = strShipperDetailsMaster;
            }
            else if (strflag == "C")
            {
                if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeCodeUsingTelphone))
                {
                    clsModule.clsMasterValuesAPR001ConsigneeCode = clsModule.clsMasterValuesAPR001ConsigneeCodeUsingTelphone;
                }
                else
                {
                    clsModule.clsMasterValuesAPR001ConsigneeCode = shipperCodePickup;//strShipperCodeMasterOrTelNoOCR; 
                }
                clsModule.clsMasterValuesAPR001ConsigneeTelePhoneNo = shipperTelNoPickup;
                clsModule.clsMasterValuesAPR001ConsigneeDetails = strShipperDetailsMaster;
            }
            else
            {
                if (!string.IsNullOrEmpty(clsModule.clsMasterValuesAPR001ConsigneeCodeUsingTelphone))
                {
                    clsModule.clsMasterValuesAPR001BillToCode = clsModule.clsMasterValuesAPR001ConsigneeCodeUsingTelphone;
                }
                else
                {
                    clsModule.clsMasterValuesAPR001BillToCode = shipperCodePickup; //strShipperCodeMasterOrTelNoOCR;
                }
                clsModule.clsMasterValuesAPR001BillToTelePhoneNo = shipperTelNoPickup;
                clsModule.clsMasterValuesAPR001BillToDetails = strShipperDetailsMaster;
            }
        }

        private DataTable GetDataFromSearchEngine(string strflag)
        {
            DataTable dtDataFromSearchEngine = new DataTable();
            try
            {
                DataTable GridCombinedAddressOCR = iQPUBLIC.Common.GetDataTable();
                string CombinedAddressOCR = string.Empty;
                string NameOCR = string.Empty;
                string AddressLineOCR = string.Empty;
                string CityOCR = string.Empty;
                string StateOCR = string.Empty;
                string ZipOCR = string.Empty;

                if (strflag == "S")
                {
                    // for shipper
                    ClsSearchEngin.isCMBTPCMBTCpresentShipper = "N";
                    ClsSearchEngin.isCMBTPCMBTCpresentValueShipper = string.Empty;
                    ClsSearchEngin.ShipperPrepaidCodeDetail = string.Empty;
                    ClsSearchEngin.ShipperCollectCodeDetail = string.Empty;

                    ClsSearchEngin.clsPickupValuesFRP001BSMatched = string.Empty;
                    ClsSearchEngin.clsPickupValuesFRP001BSMatchedCode = string.Empty;
                    ClsSearchEngin.clsPickupValuesFRP001BSMatchedDetails = string.Empty;

                    NameOCR = clsF3.ClsOCRValues.ClsOCRValuesShipperName.Trim();
                    AddressLineOCR = clsF3.ClsOCRValues.ClsOCRValuesShipperAddressLine1.Trim();
                    CityOCR = clsF3.ClsOCRValues.ClsOCRValuesShipperCity.Trim();
                    StateOCR = clsF3.ClsOCRValues.ClsOCRValuesShipperState.Trim();
                    ZipOCR = clsF3.ClsOCRValues.ClsOCRValuesShipperZip.Trim();
                    if (!string.IsNullOrEmpty(AddressLineOCR) && !string.IsNullOrEmpty(CityOCR) && !string.IsNullOrEmpty(StateOCR) && !string.IsNullOrEmpty(ZipOCR))
                    {
                        CombinedAddressOCR = AddressLineOCR.Trim() + " " + CityOCR.Trim() + " " + StateOCR.Trim() + " " + ZipOCR.Trim();
                    }
                }
                else if (strflag == "C")
                {
                    // consignee
                    ClsSearchEngin.isCMBTPCMBTCpresent = "N";
                    ClsSearchEngin.isCMBTPCMBTCpresentValue = string.Empty;
                    ClsSearchEngin.ConsigneePrepaidCodeDetail = string.Empty;
                    ClsSearchEngin.ConsigneeCollectCodeDetail = string.Empty;


                    NameOCR = clsF3.ClsOCRValues.ClsOCRValuesConsigneeName.Trim();  //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(4, GridCombinedAddressOCR));
                    AddressLineOCR = clsF3.ClsOCRValues.ClsOCRValuesConsigneeAddressLine1.Trim(); //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(5, GridCombinedAddressOCR));
                    CityOCR = clsF3.ClsOCRValues.ClsOCRValuesConsigneeCity.Trim(); //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(6, GridCombinedAddressOCR));
                    StateOCR = clsF3.ClsOCRValues.ClsOCRValuesConsigneeState.Trim();//Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(7, GridCombinedAddressOCR));
                    ZipOCR = clsF3.ClsOCRValues.ClsOCRValuesConsigneeZip.Trim();//Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(8, GridCombinedAddressOCR));
                    if (!string.IsNullOrEmpty(AddressLineOCR) && !string.IsNullOrEmpty(CityOCR) && !string.IsNullOrEmpty(StateOCR) && !string.IsNullOrEmpty(ZipOCR))
                    {
                        CombinedAddressOCR = AddressLineOCR.Trim() + " " + CityOCR.Trim() + " " + StateOCR.Trim() + " " + ZipOCR.Trim();
                    }
                }
                else
                {
                    NameOCR = clsF3.ClsOCRValues.ClsOCRValuesBillToName.Trim();  //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(4, GridCombinedAddressOCR));
                    AddressLineOCR = clsF3.ClsOCRValues.ClsOCRValuesBillToAddressLine1.Trim(); //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(5, GridCombinedAddressOCR));
                    CityOCR = clsF3.ClsOCRValues.ClsOCRValuesBillToCity.Trim(); //Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(6, GridCombinedAddressOCR));
                    StateOCR = clsF3.ClsOCRValues.ClsOCRValuesBillToState.Trim();//Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(7, GridCombinedAddressOCR));
                    ZipOCR = clsF3.ClsOCRValues.ClsOCRValuesBillToZip.Trim();//Convert.ToString(iQPUBLIC.Common.GetHeaderDataForCell(8, GridCombinedAddressOCR));
                    if (!string.IsNullOrEmpty(AddressLineOCR) && !string.IsNullOrEmpty(CityOCR) && !string.IsNullOrEmpty(StateOCR) && !string.IsNullOrEmpty(ZipOCR))
                    {
                        CombinedAddressOCR = AddressLineOCR.Trim() + " " + CityOCR.Trim() + " " + StateOCR.Trim() + " " + ZipOCR.Trim();
                    }
                }
                string TerminalID = string.Empty;
                DataTable dtKeys = new DataTable();
                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSET2_PickupShipperDetails"))
                {
                    dtKeys = (DataTable)iQPUBLIC.PublicComponents.htMyVariable["BOLSET2_PickupShipperDetails"];
                }
                if (dtKeys != null && dtKeys.Rows.Count > 0)
                {
                    if (strflag == "S")
                    {
                        TerminalID = dtKeys.Rows[0]["ShipperTerminalID"].ToString().Trim();
                    }
                    else if (strflag == "C")
                    {
                        TerminalID = dtKeys.Rows[0]["ConsigneeTerminalID"].ToString().Trim();
                    }
                    else  // TO DO as third party has no terminal id
                    {
                        TerminalID = ""; //dtKeys.Rows[0]["ShipperTerminalID"].ToString();
                    }
                }

                if (!string.IsNullOrEmpty(CombinedAddressOCR))
                {
                    clsOleDBDataAccess objOleDB = new clsOleDBDataAccess();
                    objOleDB.propConnection = new System.Data.OleDb.OleDbConnection();
                    objOleDB.propConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + mdbFileConnString + "'";
                    try
                    {
                        // Remove all special characters except space "\s" // added on 5-feb-2021  -- start
                        //CombinedAddressOCR = Regex.Replace(CombinedAddressOCR, "[^a-zA-Z0-9\\s]", "");
                        //DataTable AbbreviationTable = clsSearchCustomer.GetAbbreviationsTable(objOleDB);
                        //CombinedAddressOCR = clsSearchCustomer.ReplaceAddressAbbreviation(CombinedAddressOCR, AbbreviationTable);
                        // added on 5-feb-2021  -- end

                        AddressSection objDBAddress = clsParseUSAddress.ParseUSAddress(CombinedAddressOCR.Trim(), Module1.dtUSCityStateZip, Module1.dtUSAStateCodeName);
                        SearchParameters searchParameters = new SearchParameters();
                        if (objDBAddress.IsAddressParsed == true)
                        {
                            //searchParameters.CustomerName = NameOCR.Split(' ')[0];
                            // check length of the Name's first word on 14-Jan-2021 -- start
                            string NameFirstWord = NameOCR.Split(' ')[0];
                            string[] arrName = NameOCR.Split(' ');
                            if (arrName.Length > 0)
                            {
                                // check if it greater than 2
                                if (NameFirstWord.Trim().Length > 2)
                                {
                                    searchParameters.CustomerName = arrName[0];
                                }
                                else
                                {
                                    if (NameFirstWord.Length <= 2 && arrName.Length >= 2)
                                    {
                                        searchParameters.CustomerName = arrName[0] + " " + arrName[1];
                                    }
                                }

                            }

                            // Remove all starting special characters from the string
                            searchParameters.CustomerName = Regex.Replace(searchParameters.CustomerName, "^[^a-zA-Z0-9]+", "");
                            // Remove all ending special characters from the string
                            searchParameters.CustomerName = Regex.Replace(searchParameters.CustomerName, "[^a-zA-Z0-9]*$", "");

                            // send Customer Name OBL 
                            // Remove all starting special characters from the string
                            searchParameters.CustomerNameOBL = NameOCR;
                            searchParameters.CustomerNameOBL = Regex.Replace(searchParameters.CustomerNameOBL, "^[^a-zA-Z0-9]+", "");
                            // Remove all ending special characters from the string
                            searchParameters.CustomerNameOBL = Regex.Replace(searchParameters.CustomerNameOBL, "[^a-zA-Z0-9]*$", "");

                            //check length of the Name's first word on 14-Jan-2021 -- end
                            searchParameters.TerminalID = TerminalID;
                            searchParameters.StreetNumber = objDBAddress.StreetNumber;
                            searchParameters.DirectionFirstLetter = "";
                            searchParameters.ZipCode = objDBAddress.ZipCode;
                            searchParameters.CustomerFirstLetter = "";
                            searchParameters.City = objDBAddress.City;
                            SearchParameters.SearchType searchType = new SearchParameters.SearchType();
                            SearchParameters.CustomerType customerType = new SearchParameters.CustomerType();
                            // Search Type =>  check for BS
                            if (strflag == "S" && clsModule.clsPickupValuesFRP001BS == "Y")
                            {
                                searchType = SearchParameters.SearchType.BlindShipper;
                            }
                            else
                            {
                                searchType = SearchParameters.SearchType.Normal;
                            }
                            // Customer Type 
                            if (strflag == "S")
                            {
                                customerType = SearchParameters.CustomerType.Shipper;

                            }
                            else if (strflag == "C")
                            {
                                customerType = SearchParameters.CustomerType.Consignee;
                            }
                            else
                            {
                                customerType = SearchParameters.CustomerType.ThirdParty;
                            }

                            #region "Use Connection Object"
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
                            dtDataFromSearchEngine = clsSearchCustomer.SearchAndFetchMatchingCustomers(objDBAddress, searchParameters, connection, objOleDB, customerType, searchType);
                            if (dtDataFromSearchEngine != null)
                            {
                                if (dtDataFromSearchEngine.Rows.Count > 0)
                                {
                                    if (strflag == "S")
                                    {
                                        ClsSearchEngin.ClsSearchEnginShipperCode = dtDataFromSearchEngine.Rows[0]["CMCUST"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginShipperName = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim();
                                        if (dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() == dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim())
                                        {
                                            ClsSearchEngin.ClsSearchEnginShipperAddressLine1 = dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim();
                                        }
                                        else
                                        {
                                            ClsSearchEngin.ClsSearchEnginShipperAddressLine1 = dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + " " + dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim();
                                        }
                                        ClsSearchEngin.ClsSearchEnginShipperCity = dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginShipperState = dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginShipperZip = dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginShipperTelePhoneNo = dtDataFromSearchEngine.Rows[0]["CMPHON"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginShipperDetailsWithPipe = dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchCommonStatusFlagShipper = ClsSearchEngin.ClsSearchCommonStatusFlag;
                                        // 1) check the shipper code is non miscellaneous code  2) Shipper is attached with the Bill-To 
                                        #region commented code    
                                        //if (!ClsSearchEngin.ClsSearchEnginShipperCode.StartsWith("99"))
                                        //{
                                        //    ClsSearchEngin.isCMBTPCMBTCpresentShipper = "N";
                                        //    ClsSearchEngin.isCMBTPCMBTCpresentValueShipper = string.Empty;

                                        //    string CMBTP = string.Empty;
                                        //    string CMBTC = string.Empty;
                                        //    CMBTP = dtDataFromSearchEngine.Rows[0]["CMBTP"].ToString().Trim();
                                        //    CMBTC = dtDataFromSearchEngine.Rows[0]["CMBTC"].ToString().Trim();
                                        //    if (string.IsNullOrEmpty(CMBTP) && string.IsNullOrEmpty(CMBTC))
                                        //    {
                                        //        ClsSearchEngin.isCMBTPCMBTCpresentShipper = "N";
                                        //        ClsSearchEngin.isCMBTPCMBTCpresentValueShipper = CMBTP + "~" + CMBTC;
                                        //    }
                                        //    if (!string.IsNullOrEmpty(CMBTP) || !string.IsNullOrEmpty(CMBTC))
                                        //    {
                                        //        ClsSearchEngin.isCMBTPCMBTCpresentShipper = "Y";
                                        //        ClsSearchEngin.isCMBTPCMBTCpresentValueShipper = CMBTP + "~" + CMBTC;
                                        //        if (dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() == dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim())
                                        //        {
                                        //            ClsSearchEngin.ShipperPrepaidCodeDetail = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        //            ClsSearchEngin.ShipperCollectCodeDetail = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        //        }
                                        //        else
                                        //        {
                                        //            ClsSearchEngin.ShipperPrepaidCodeDetail = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + " " + dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        //            ClsSearchEngin.ShipperCollectCodeDetail = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + " " + dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        //        }

                                        //    }
                                        //    // Check for BS
                                        //    //if (clsModule.clsPickupValuesFRP001BS == "Y")
                                        //    //{
                                        //    //    ClsSearchEngin.clsPickupValuesFRP001BSMatched = "Y";
                                        //    //    ClsSearchEngin.clsPickupValuesFRP001BSMatchedCode = "BS - " + dtDataFromSearchEngine.Rows[0]["CMCUST"].ToString().Trim();
                                        //    //    ClsSearchEngin.clsPickupValuesFRP001BSMatchedDetails = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + " " + dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        //    //}
                                        //}
                                        #endregion
                                    }
                                    else if (strflag == "C")
                                    {
                                        ClsSearchEngin.ClsSearchEnginConsigneeCode = dtDataFromSearchEngine.Rows[0]["CMCUST"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginConsigneeName = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim();
                                        if (dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() == dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim())
                                        {
                                            ClsSearchEngin.ClsSearchEnginConsigneeAddressLine1 = dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim();
                                        }
                                        else
                                        {
                                            ClsSearchEngin.ClsSearchEnginConsigneeAddressLine1 = dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + " " + dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim();
                                        }
                                        ClsSearchEngin.ClsSearchEnginConsigneeCity = dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginConsigneeState = dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginConsigneeZip = dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginConsigneeTelePhoneNo = dtDataFromSearchEngine.Rows[0]["CMPHON"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginConsigneeDetailsWithPipe = dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchCommonStatusFlagConsignee = ClsSearchEngin.ClsSearchCommonStatusFlag;
                                        // 1) Check the consignee code is non miscellaneous code 2) Consignee is attached with Bill-To
                                        #region commented code
                                        //if (!ClsSearchEngin.ClsSearchEnginConsigneeCode.StartsWith("98"))
                                        //{
                                        //    ClsSearchEngin.isCMBTPCMBTCpresent = "N";
                                        //    ClsSearchEngin.isCMBTPCMBTCpresentValue = string.Empty;
                                        //    string CMBTP = string.Empty;
                                        //    string CMBTC = string.Empty;
                                        //    CMBTP = dtDataFromSearchEngine.Rows[0]["CMBTP"].ToString().Trim();
                                        //    CMBTC = dtDataFromSearchEngine.Rows[0]["CMBTC"].ToString().Trim();
                                        //    if (string.IsNullOrEmpty(CMBTP) && string.IsNullOrEmpty(CMBTC))
                                        //    {
                                        //        ClsSearchEngin.isCMBTPCMBTCpresent = "N";
                                        //        ClsSearchEngin.isCMBTPCMBTCpresentValue = CMBTP + "~" + CMBTC;
                                        //    }
                                        //    if (!string.IsNullOrEmpty(CMBTP) || !string.IsNullOrEmpty(CMBTC))
                                        //    {
                                        //        ClsSearchEngin.isCMBTPCMBTCpresent = "Y";
                                        //        ClsSearchEngin.isCMBTPCMBTCpresentValue = CMBTP + "~" + CMBTC;
                                        //        if (dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() == dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim())
                                        //        {
                                        //            ClsSearchEngin.ConsigneePrepaidCodeDetail = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        //            ClsSearchEngin.ConsigneeCollectCodeDetail = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        //        }
                                        //        else
                                        //        {
                                        //            ClsSearchEngin.ConsigneePrepaidCodeDetail = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + " " + dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        //            ClsSearchEngin.ConsigneeCollectCodeDetail = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + " " + dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim() + "~" + dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        //        }

                                        //    }
                                        //}
                                        #endregion
                                    }
                                    else
                                    {
                                        ClsSearchEngin.ClsSearchEnginBillToCode = dtDataFromSearchEngine.Rows[0]["CMCUST"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginBillToName = dtDataFromSearchEngine.Rows[0]["CMNAME"].ToString().Trim();
                                        if (dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() == dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim())
                                        {
                                            ClsSearchEngin.ClsSearchEnginBillToAddressLine1 = dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim();
                                        }
                                        else
                                        {
                                            ClsSearchEngin.ClsSearchEnginBillToAddressLine1 = dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + " " + dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim();
                                        }
                                        ClsSearchEngin.ClsSearchEnginBillToCity = dtDataFromSearchEngine.Rows[0]["CMCITY"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginBillToState = dtDataFromSearchEngine.Rows[0]["CMST"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginBillToZip = dtDataFromSearchEngine.Rows[0]["CMZIP"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginBillToTelePhoneNo = dtDataFromSearchEngine.Rows[0]["CMPHON"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchEnginBillToDetailsWithPipe = dtDataFromSearchEngine.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtDataFromSearchEngine.Rows[0]["CMADR2"].ToString().Trim();
                                        ClsSearchEngin.ClsSearchCommonStatusFlagBillTo = ClsSearchEngin.ClsSearchCommonStatusFlag;
                                    }
                                }
                                // check if first word of name lenght is equal to 2 then again serach with adding space in between to letters
                                //else
                                //{
                                //    //if (NameFirstWord.Length > 0 && NameFirstWord.Trim().Length == 2 && arrName.Length>0)
                                //{
                                //char[] letters = arrName[0].Trim().ToCharArray();
                                //searchParameters.CustomerName = letters[0] + " " + letters[1];

                                //}
                                //}
                            }
                        }
                        else
                        {
                            dtDataFromSearchEngine = null;
                        }
                    }

                    catch (Exception ex)
                    {
                        //MessageBox.Show(string.Concat("Exception -> ", ex.Message), "Parse Address");
                    }
                }
            }
            catch (Exception ex)
            {

                //MessageBox.Show(Ex.Message);
            }
            finally
            {

            }

            return dtDataFromSearchEngine;

        }

        public void GetMailToAttachedData(string customercode, string flag)
        {
            try
            {

                string queryStringMaster = "select CMMAD1,CMMAD2,CMMCIT,CMMST,CMMZIP from ARP001 where CMSTAT = 'A' AND CMCUST='" + customercode + "'";
                DataTable dtMaster_Shipper_Details = GetDataTableByText(As400_ConnectionString, queryStringMaster);
                if (dtMaster_Shipper_Details != null && dtMaster_Shipper_Details.Rows.Count > 0)
                {
                    if (flag == "S")
                    {

                        if (!string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMCIT"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMST"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMZIP"].ToString().Trim()))
                        {
                            clsModule.clsMasterValuesAPR001MailToAttachedShipperDetail = dtMaster_Shipper_Details.Rows[0]["CMMAD1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMAD2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMCIT"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMZIP"].ToString().Trim();
                            // address1 and address2 with pipe
                            clsModule.clsMasterValuesAPR001MailToAttachedShipperWithPipe = dtMaster_Shipper_Details.Rows[0]["CMMAD1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMMAD2"].ToString().Trim();
                        }


                    }
                    else if (flag == "C")
                    {
                        if (!string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMCIT"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMST"].ToString().Trim()) && !string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMMZIP"].ToString().Trim()))
                        {
                            clsModule.clsMasterValuesAPR001MailToAttachedConsigneeDetail = dtMaster_Shipper_Details.Rows[0]["CMMAD1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMAD2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMCIT"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMMZIP"].ToString().Trim();
                            // address1 and address2 with pipe
                            clsModule.clsMasterValuesAPR001MailToAttachedConsigneeWithPipe = dtMaster_Shipper_Details.Rows[0]["CMMAD1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMMAD2"].ToString().Trim();
                        }

                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void GetBillToAttachedData(string customercode, string flag)
        {
            try
            {
                string queryStringMaster = "select CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP,CMBTP,CMBTC from ARP001 where CMSTAT = 'A' AND CMCUST='" + customercode + "'";
                DataTable dtMaster_Shipper_Details = GetDataTableByText(As400_ConnectionString, queryStringMaster);
                if (dtMaster_Shipper_Details != null && dtMaster_Shipper_Details.Rows.Count > 0)
                {
                    if (flag == "S")
                    {
                        if (!string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMBTP"].ToString().Trim()))
                        {
                            clsModule.isCMBTPCMBTCpresentShipper = "Y";
                            clsModule.isCMBTPCMBTCpresentValueShipper = dtMaster_Shipper_Details.Rows[0]["CMBTP"].ToString().Trim();
                            //if (dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim())
                            //{
                            //    clsModule.ShipperPrepaidCodeDetail = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                            //}
                            //else
                            //{
                            //    clsModule.ShipperPrepaidCodeDetail = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                            //}
                        }

                    }
                    else if (flag == "C")
                    {
                        if (!string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMBTC"].ToString().Trim()))
                        {
                            clsModule.isCMBTPCMBTCpresent = "Y";
                            clsModule.isCMBTPCMBTCpresentValue = dtMaster_Shipper_Details.Rows[0]["CMBTC"].ToString().Trim();
                            //if (dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim())
                            //{
                            //    clsModule.ConsigneeCollectCodeDetail = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                            //}
                            //else
                            //{
                            //    clsModule.ConsigneeCollectCodeDetail = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                            //}
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        public void GetBillToAttachedDataDetail(string customercode, string flag)
        {
            try
            {
                string queryStringMaster = "select CMNAME,CMADR1,CMADR2,CMCITY,CMST,CMZIP,CMBTP,CMBTC from ARP001 where CMSTAT = 'A' AND CMCUST='" + customercode + "'";
                DataTable dtMaster_Shipper_Details = GetDataTableByText(As400_ConnectionString, queryStringMaster);
                if (dtMaster_Shipper_Details != null && dtMaster_Shipper_Details.Rows.Count > 0)
                {
                    if (flag == "S")
                    {
                        //if (!string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMBTP"].ToString().Trim()))
                        //{
                        //clsModule.isCMBTPCMBTCpresentShipper = "Y";
                        //clsModule.isCMBTPCMBTCpresentValueShipper = dtMaster_Shipper_Details.Rows[0]["CMBTP"].ToString().Trim();
                        if (dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim())
                        {
                            clsModule.ShipperPrepaidCodeDetail = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();

                        }
                        else
                        {
                            clsModule.ShipperPrepaidCodeDetail = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                        }
                        clsModule.clsMasterValuesAPR001BillToAttachedShipperWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();

                        //}

                    }
                    else if (flag == "C")
                    {
                        //if (!string.IsNullOrEmpty(dtMaster_Shipper_Details.Rows[0]["CMBTC"].ToString().Trim()))
                        //{
                        //clsModule.isCMBTPCMBTCpresent = "Y";
                        //clsModule.isCMBTPCMBTCpresentValue = dtMaster_Shipper_Details.Rows[0]["CMBTC"].ToString().Trim();
                        if (dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() == dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim())
                        {
                            clsModule.ConsigneeCollectCodeDetail = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                        }
                        else
                        {
                            clsModule.ConsigneeCollectCodeDetail = dtMaster_Shipper_Details.Rows[0]["CMNAME"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + " " + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMCITY"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMST"].ToString().Trim() + "~" + dtMaster_Shipper_Details.Rows[0]["CMZIP"].ToString().Trim();
                        }
                        clsModule.clsMasterValuesAPR001BillToAttachedConsigneeWithPipe = dtMaster_Shipper_Details.Rows[0]["CMADR1"].ToString().Trim() + "||" + dtMaster_Shipper_Details.Rows[0]["CMADR2"].ToString().Trim();
                        //}
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        public class ClsSearchEngin
        {
            // shipper
            public static string ClsSearchEnginShipperCode { get; set; }
            public static string ClsSearchEnginShipperName { get; set; }
            public static string ClsSearchEnginShipperAddressLine1 { get; set; }
            public static string ClsSearchEnginShipperCity { get; set; }
            public static string ClsSearchEnginShipperState { get; set; }
            public static string ClsSearchEnginShipperZip { get; set; }
            public static string ClsSearchEnginShipperTelePhoneNo { get; set; }
            //consignee
            public static string ClsSearchEnginConsigneeCode { get; set; }
            public static string ClsSearchEnginConsigneeName { get; set; }
            public static string ClsSearchEnginConsigneeAddressLine1 { get; set; }
            public static string ClsSearchEnginConsigneeCity { get; set; }
            public static string ClsSearchEnginConsigneeState { get; set; }
            public static string ClsSearchEnginConsigneeZip { get; set; }
            public static string ClsSearchEnginConsigneeTelePhoneNo { get; set; }
            // Bill To 3rd party
            public static string ClsSearchEnginBillToCode { get; set; }
            public static string ClsSearchEnginBillToName { get; set; }
            public static string ClsSearchEnginBillToAddressLine1 { get; set; }
            public static string ClsSearchEnginBillToCity { get; set; }
            public static string ClsSearchEnginBillToState { get; set; }
            public static string ClsSearchEnginBillToZip { get; set; }
            public static string ClsSearchEnginBillToTelePhoneNo { get; set; }

            // Other => 
            public static string isCMBTPCMBTCpresent { get; set; }
            public static string isCMBTPCMBTCpresentValue { get; set; }

            public static string isCMBTPCMBTCpresentShipper { get; set; }
            public static string isCMBTPCMBTCpresentValueShipper { get; set; }

            public static string ShipperPrepaidCodeDetail { get; set; }
            public static string ShipperCollectCodeDetail { get; set; }
            public static string ConsigneePrepaidCodeDetail { get; set; }
            public static string ConsigneeCollectCodeDetail { get; set; }

            public static string clsPickupValuesFRP001BSMatched { get; set; }
            public static string clsPickupValuesFRP001BSMatchedCode { get; set; }
            public static string clsPickupValuesFRP001BSMatchedDetails { get; set; }

            //Address1 and Address2 with pipe 
            public static string ClsSearchEnginShipperDetailsWithPipe { get; set; }
            public static string ClsSearchEnginConsigneeDetailsWithPipe { get; set; }
            public static string ClsSearchEnginBillToDetailsWithPipe { get; set; }

            // alise case
            public static string ClsSearchEnginAliseShipperNameARP033 { get; set; }
            public static string ClsSearchEnginAliseConsigneeNameARP033 { get; set; }
            public static string ClsSearchEnginAliseBillToNameARP033 { get; set; }

            public static string ClsSearchCommonStatusFlag { get; set; }
            public static string ClsSearchCommonStatusFlagShipper { get; set; }
            public static string ClsSearchCommonStatusFlagConsignee { get; set; }
            public static string ClsSearchCommonStatusFlagBillTo { get; set; }

            //public static string ClsSearchEngine
        }
        public class clsModule
        {
            //  Pick up details => Shipper
            public static string clsPickupValuesFRP001ShipperCode { get; set; }
            public static string clsPickupValuesFRP001ShipperDetailsPickUp { get; set; }
            public static string clsPickupValuesFRP001ShipperName { get; set; }
            public static string clsPickupValuesFRP001ShipperAddress1 { get; set; }
            public static string clsPickupValuesFRP001ShipperCity { get; set; }
            public static string clsPickupValuesFRP001ShipperState { get; set; }
            public static string clsPickupValuesFRP001ShipperZip { get; set; }

            public static string clsPickupValuesFRP001RAT { get; set; }
            public static string clsPickupValuesFRP001ShipperTerminalID { get; set; }

            public static string clsPickupValuesFRP001BS { get; set; }
            public static string clsPickupValuesFRP001BSMatched { get; set; }

            //  Pick up details => Consignee
            public static string clsPickupValuesFRP001ConsigneeCode { get; set; }
            public static string clsPickupValuesFRP001ConsigneeName { get; set; }
            public static string clsPickupValuesFRP001ConsigneeAddress1 { get; set; }
            public static string clsPickupValuesFRP001ConsigneeCity { get; set; }
            public static string clsPickupValuesFRP001ConsigneeState { get; set; }
            public static string clsPickupValuesFRP001ConsigneeZip { get; set; }
            //public static string clsPickupValuesFRP001ConsigneeTelePhone { get; set; }

            public static string clsPickupValuesFRP001HAT { get; set; }
            public static string clsPickupValuesFRP001ConsigneeTerminalID { get; set; }

            // Bill To  3rd Party
            public static string clsPickupValuesFRP001BillToCode { get; set; }
            public static string clsPickupValuesFRP001BillToName { get; set; }
            public static string clsPickupValuesFRP001BillToAddress1 { get; set; }
            public static string clsPickupValuesFRP001BillToCity { get; set; }
            public static string clsPickupValuesFRP001BillToState { get; set; }
            public static string clsPickupValuesFRP001BillToZip { get; set; }


            // Master details => 

            // Shipper 
            public static string clsMasterValuesAPR001ShipperCode { get; set; }
            public static string clsMasterValuesAPR001ShipperTelePhoneNo { get; set; }
            public static string clsMasterValuesAPR001ShipperDetails { get; set; }

            public static string clsMasterValuesAPR001ShipperRATCode { get; set; }
            public static string clsMasterValuesAPR001ShipperRATCodeFinal { get; set; }
            public static string clsMasterValuesAPR001ShipperRATDetails { get; set; }
            public static string clsMasterValuesAPR001ShipperRATTelePhoneNo { get; set; }


            // Consignee  
            public static string clsMasterValuesAPR001ConsigneeCode { get; set; }

            public static string clsMasterValuesAPR001ConsigneeTelePhoneNo { get; set; }
            public static string clsMasterValuesAPR001ConsigneeDetails { get; set; }

            public static string clsMasterValuesAPR001ConsigneeHATCode { get; set; }
            public static string clsMasterValuesAPR001ConsigneeHATCodeFinal { get; set; }
            public static string clsMasterValuesAPR001ConsigneeHATDetails { get; set; }
            public static string clsMasterValuesAPR001ConsigneeHATTelePhoneNo { get; set; }

            // BillTo 3rd party
            public static string clsMasterValuesAPR001BillToCode { get; set; }
            public static string clsMasterValuesAPR001BillToTelePhoneNo { get; set; }
            public static string clsMasterValuesAPR001BillToDetails { get; set; }

            // Master or Search Enginee  =>
            // Shipper
            public static string OCR_Matched_MasterDB_KeyShipper { get; set; }
            public static string OCR_matched_SearchenagineShipper { get; set; }
            // Consignee
            public static string OCR_Matched_MasterDB_KeyConsignee { get; set; }
            public static string OCR_matched_SearchenagineConsignee { get; set; }

            // Bill To 3rd party
            public static string OCR_Matched_MasterDB_KeyBillTo { get; set; }
            public static string OCR_matched_SearchenagineBillTo { get; set; }

            // Othere =>

            public static string isCMBTPCMBTCpresent { get; set; }
            public static string isCMBTPCMBTCpresentValue { get; set; }

            public static string ConsigneeCodeFinal { get; set; }

            public static string isCMBTPCMBTCpresentShipper { get; set; }
            public static string isCMBTPCMBTCpresentValueShipper { get; set; }

            public static string isBillToAttachedShipper { get; set; }
            public static string isBillToAttachedConsignee { get; set; }
            public static string ShipperCodeFinal { get; set; }


            public static string ShipperPrepaidCodeDetail { get; set; }
            public static string ShipperCollectCodeDetail { get; set; }
            public static string ConsigneePrepaidCodeDetail { get; set; }
            public static string ConsigneeCollectCodeDetail { get; set; }


            public static string clsPickupValuesFRP001BSMatchedCode { get; set; }
            public static string clsPickupValuesFRP001BSMatchedDetails { get; set; }

            // Final Values => 
            // shipper => 
            public static string clsFinalValuesShipperName { get; set; }
            public static string clsFinalValuesShipperAddress1 { get; set; }
            public static string clsFinalValuesShipperCity { get; set; }
            public static string clsFinalValuesShipperState { get; set; }
            public static string clsFinalValuesShipperZip { get; set; }
            // Consignee
            public static string clsFinalValuesConsigneeName { get; set; }
            public static string clsFinalValuesConsigneeAddress1 { get; set; }
            public static string clsFinalValuesConsigneeCity { get; set; }
            public static string clsFinalValuesConsigneeState { get; set; }
            public static string clsFinalValuesConsigneeZip { get; set; }

            // Bill To
            public static string clsFinalValuesBillToName { get; set; }
            public static string clsFinalValuesBillToAddress1 { get; set; }
            public static string clsFinalValuesBillToCity { get; set; }
            public static string clsFinalValuesBillToState { get; set; }
            public static string clsFinalValuesBillToZip { get; set; }

            // Mail To attached 
            public static string clsMasterValuesAPR001MailToAttachedShipperDetail { get; set; }
            public static string clsMasterValuesAPR001MailToAttachedConsigneeDetail { get; set; }

            public static string isMailToAttachedShipper { get; set; }
            public static string isMailToAttachedConsignee { get; set; }

            // Address 1 and Address 2 with pipe separted 

            // shipper
            public static string clsMasterValuesAPR001ShipperDetailsWithPipe { get; set; }
            //public static string clsMasterValuesAPR001ShipperRATDetailsWithPipe { get; set; }
            //public static string clsMasterValuesAPR001ShipperBSDetailsWithPipe { get; set; }

            // consignee
            public static string clsMasterValuesAPR001ConsigneeDetailsWithPipe { get; set; }
            //public static string clsMasterValuesAPR001ConsigneeHATDetailsWithPipe { get; set; }

            // 3rd party
            public static string clsMasterValuesAPR001MailToAttachedShipperWithPipe { get; set; }
            public static string clsMasterValuesAPR001MailToAttachedConsigneeWithPipe { get; set; }
            public static string clsMasterValuesAPR001BillToAttachedShipperWithPipe { get; set; }
            public static string clsMasterValuesAPR001BillToAttachedConsigneeWithPipe { get; set; }

            public static string clsMasterValuesAPR001AliseShipperName { get; set; }
            public static string clsMasterValuesAPR001AliseConsigneeName { get; set; }
            public static string clsMasterValuesAPR001AliseBillToName { get; set; }

            // Search from telephone number
            public static string clsMasterValuesAPR001ShipperCodeUsingTelphone { get; set; }
            public static string clsMasterValuesAPR001ConsigneeCodeUsingTelphone { get; set; }
            public static string clsMasterValuesAPR001BillToCodeUsingTelphone { get; set; }

            // check valid code
            public static string isShipperValidFlag { get; set; }
            public static string isConsigneeValidFlag { get; set; }

            // find consignee miscellaneous code
            public static string isConsigneeCodeRedFlag { get; set; }
            public static string isConsigneeAllGreenFlag { get; set; }

            // Added on 13-July-2021
            public static string ShipperKMFinalCode { get; set; }
            public static string ShipperKMFinalName { get; set; }
            public static string ShipperKMFinalAddressLine { get; set; }
            public static string ShipperKMFinalCity { get; set; }
            public static string ShipperKMFinalState { get; set; }
            public static string ShipperKMFinalZip { get; set; }

            public static string ShipperKMFlag { get; set; }

            public static string ShipperKMOBLName { get; set; }
            public static string ShipperKMOBLAddressLine { get; set; }
           
        }
        public static DataTable GetDataTableByText(string connectionString, string queryString, OdbcParameter[] parameters = null)
        //public static DataTable GetDataTableByText(string queryString, OdbcParameter[] parameters = null)
        {
            DataTable dataTable = new DataTable();
            try
            {
                //using (OdbcConnection connection = new OdbcConnection(connectionString))
                //{
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
                    db2Connection.ConnectionString = connectionString; //iQPUBLIC.PublicComponents.DbConnectionString1;

                    if (db2Connection.State != ConnectionState.Open)
                    {
                        db2Connection.Open();

                        iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", db2Connection);

                    }

                }

                OdbcConnection connection = (OdbcConnection)iQPUBLIC.PublicComponents.htMyVariable["DB2Conn"];
                if (connection.State != ConnectionState.Open)
                {
                    //OdbcConnection db2Connection = new OdbcConnection();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_DB2Conn") == true)
                    {
                        connectionString = (string)iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_DB2Conn"];
                    }
                    else
                    {
                        connectionString = ConfigurationManager.ConnectionStrings["AS400DBConnString"].ToString();

                    }
                    connection.ConnectionString = connectionString; // iQPUBLIC.PublicComponents.DbConnectionString1; 
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
                //}
                return dataTable;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                if (dataTable != null)
                {
                    dataTable.Dispose();
                }
            }
        }

        //public string MatchAddress(string OCRAddress, string MasterAddress, clsOleDBDataAccess objOleDB, string Flag)
        public string MatchAddress(string OCRAddress, string MasterAddress, string Flag)
        {
            string MatchAddress = string.Empty;
            try
            {
                AddressSection AddressSectionsFromOCR = clsParseUSAddress.ParseUSAddress(OCRAddress.Trim(), Module1.dtUSCityStateZip, Module1.dtUSAStateCodeName);
                AddressSection objDBAddress = clsParseUSAddress.ParseUSAddress(MasterAddress.Trim(), Module1.dtUSCityStateZip, Module1.dtUSAStateCodeName);
                SearchParameters.CustomerType customerType = new SearchParameters.CustomerType();
                if (AddressSectionsFromOCR.IsAddressParsed == true && objDBAddress.IsAddressParsed == true)
                {
                    if (Flag == "S")
                    {
                        customerType = SearchParameters.CustomerType.Shipper;
                    }
                    else if (Flag == "C")
                    {
                        customerType = SearchParameters.CustomerType.Consignee;
                    }
                    else
                    {
                        customerType = SearchParameters.CustomerType.ThirdParty;
                    }

                    // DataTable AbbreviationTable = clsSearchCustomer.GetAbbreviationsTable(objOleDB);
                    // For OCR Address Line
                    //string ocrValue = clsSearchCustomer.ReplaceAddressAbbreviation(AddressSectionsFromOCR.AddressLine1.Trim(), AbbreviationTable); // Commented and added below to handle the exception
                    string ocrValue = string.Empty;
                    if (!string.IsNullOrEmpty(AddressSectionsFromOCR.AddressLine1))
                    {
                        ocrValue = clsSearchCustomer.ReplaceAddressAbbreviation(AddressSectionsFromOCR.AddressLine1.Trim(), Module1.dtAbbreviationTable);
                    }
                    AddressSectionsFromOCR.AddressLine1 = ocrValue;
                    // For DB Address Line
                    ProcessAddressParameters objProcessAddressParameters = new ProcessAddressParameters();
                    objProcessAddressParameters.AddressLine1 = objDBAddress.AddressLine1;
                    objProcessAddressParameters.AddressLine2 = "";
                    objProcessAddressParameters.StreetNumber = objDBAddress.StreetNumber;
                    objProcessAddressParameters.StreetDirection = objDBAddress.Direction;

                    objDBAddress.AddressLine1 = clsSearchCustomer.ProcessAddressBeforeComparison(objProcessAddressParameters, Module1.dtAbbreviationTable); // process address line of DB and then pass it as parameter

                    string matchResults = clsSearchCustomer.MatchAddressUsingRules(objDBAddress, AddressSectionsFromOCR, customerType, Module1.dtAbbreviationTable);

                    if (!string.IsNullOrEmpty(matchResults))
                    {
                        if (matchResults.Split('~')[0] == "Y")
                        {
                            MatchAddress = "Y";
                        }
                        else
                        {
                            //MatchAddress = "N"; // commeted on 11-March-2021 
                            // if address not match then 1) Take original addresses of DB and OCR 2) Expand Abbreviation 3) Remove special character and space 3) Match  100%
                            OCRAddress = OCRAddress.Trim().ToUpper();
                            MasterAddress = MasterAddress.Trim().ToUpper();
                            OCRAddress = clsSearchCustomer.ReplaceAddressAbbreviation(OCRAddress, Module1.dtAbbreviationTable);
                            OCRAddress = Module1.func_RemoveSpecialCharacter(OCRAddress); // Removed special characters and space
                            MasterAddress = clsSearchCustomer.ReplaceAddressAbbreviation(MasterAddress, Module1.dtAbbreviationTable);
                            MasterAddress = Module1.func_RemoveSpecialCharacter(MasterAddress); // Removed special characters and space
                            clsCNCSBR clsCNCSBR = new clsCNCSBR();
                            RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(MasterAddress, OCRAddress);

                            if (retStruct.PercentageMatch == 100)
                            {
                                MatchAddress = "Y";
                            }
                            else
                            {
                                MatchAddress = "N";
                            }

                        }
                    }

                }
                else
                { MatchAddress = "N"; }

                return MatchAddress;
            }
            catch (Exception ex)
            {
                throw;
            }

        }

        public string WordMatchFromNameOld(string DBName, string OCRName)
        {
            string IsWordMatchFromName = string.Empty;
            DBName = Module1.func_RemoveSpecialWordsFromName(DBName);
            OCRName = Module1.func_RemoveSpecialWordsFromName(OCRName);
            DBName = Module1.func_RemoveSpecialCharacterwithSpace(DBName);
            OCRName = Module1.func_RemoveSpecialCharacterwithSpace(OCRName);
            string[] strDBName = DBName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            // Dictionary<string, string> DBDictionary = strDBName.Select(item => item.Split(' ')).ToDictionary(s => s[0], s => s[0]);
            Dictionary<string, string> DBDictionary = new Dictionary<string, string>();
            for (int i = 0; i < strDBName.Length; i++)
            {
                if (DBDictionary.ContainsKey(strDBName[i]) == false)
                {
                    DBDictionary.Add(strDBName[i], strDBName[i]);
                }
            }
            string[] ocrWords = OCRName.Split(' ');
            bool isWordFound = false;
            if (ocrWords.Length > 0 && ocrWords.Length > 0)
            {
                for (int wordIndex1 = 0; wordIndex1 < ocrWords.Length; wordIndex1++)
                {
                    if (DBDictionary.ContainsKey(ocrWords[wordIndex1]))
                    {
                        if (ocrWords[wordIndex1].Length >= 3 && ocrWords[wordIndex1] != "C/O")
                        {
                            isWordFound = true;
                            break;
                        }
                    }
                    if (isWordFound)
                        break;
                }
                // if word lenght less than 3 then add space in each character
                if (isWordFound == false && ocrWords[0].Length == 3)
                {
                    char[] letters = ocrWords[0].ToUpper().Trim().ToCharArray();
                    string CustomerNameFormattedThreeLetter = letters[0] + " " + letters[1] + " " + letters[2];
                    if (DBName.Contains(CustomerNameFormattedThreeLetter))
                    {
                        isWordFound = true;
                    }
                }
                else if (isWordFound == false && ocrWords[0].Length == 2)
                {
                    char[] letters = ocrWords[0].ToUpper().Trim().ToCharArray();
                    string CustomerNameFormattedThreeLetter = letters[0] + " " + letters[1];
                    if (DBName.Contains(CustomerNameFormattedThreeLetter))
                    {
                        isWordFound = true;
                    }
                }
                else if (isWordFound == false && ocrWords[0].Contains("-"))
                {
                    string splittedCustomerNameWithoutHyphen = string.Concat(ocrWords[0].Split('-')[0], " ", ocrWords[0].Split('-')[1]);
                    if (DBName.Contains(splittedCustomerNameWithoutHyphen))
                    {
                        isWordFound = true;
                    }
                }

            }
            if (isWordFound == true)
            {
                IsWordMatchFromName = "Y";
            }
            return IsWordMatchFromName;
        }
        public int WordMatchFromName(string DBName, string OCRName)
        {
            int matchcount = 0;
            string IsWordMatchFromName = string.Empty;
            DBName = Module1.func_RemoveSpecialWordsFromName(DBName);
            OCRName = Module1.func_RemoveSpecialWordsFromName(OCRName);
            DBName = Module1.func_RemoveSpecialCharacterwithSpace(DBName);
            OCRName = Module1.func_RemoveSpecialCharacterwithSpace(OCRName);
            string[] strDBName = DBName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            // Dictionary<string, string> DBDictionary = strDBName.Select(item => item.Split(' ')).ToDictionary(s => s[0], s => s[0]);
            Dictionary<string, string> DBDictionary = new Dictionary<string, string>();
            for (int i = 0; i < strDBName.Length; i++)
            {
                if (DBDictionary.ContainsKey(strDBName[i]) == false)
                {
                    DBDictionary.Add(strDBName[i], strDBName[i]);
                }
            }
            string[] ocrWords = OCRName.Split(' ');
            bool isWordFound = false;
            if (ocrWords.Length > 0 && ocrWords.Length > 0)
            {
                for (int wordIndex1 = 0; wordIndex1 < ocrWords.Length; wordIndex1++)
                {
                    if (DBDictionary.ContainsKey(ocrWords[wordIndex1]))
                    {
                        if (ocrWords[wordIndex1].Length >= 3 && ocrWords[wordIndex1] != "C/O")
                        {
                            isWordFound = true;
                            //break;
                            matchcount++;
                        }
                    }
                    //if (isWordFound)
                    //    break;
                }
                // if word lenght less than 3 then add space in each character  
                if (isWordFound == false && ocrWords[0].Length == 3)
                {
                    char[] letters = ocrWords[0].ToUpper().Trim().ToCharArray();
                    string CustomerNameFormattedThreeLetter = letters[0] + " " + letters[1] + " " + letters[2];
                    if (DBName.Contains(CustomerNameFormattedThreeLetter))
                    {
                        //isWordFound = true;
                        matchcount++;
                    }
                }
                else if (isWordFound == false && ocrWords[0].Length == 2)
                {
                    char[] letters = ocrWords[0].ToUpper().Trim().ToCharArray();
                    string CustomerNameFormattedThreeLetter = letters[0] + " " + letters[1];
                    if (DBName.Contains(CustomerNameFormattedThreeLetter))
                    {
                        //isWordFound = true;
                        matchcount++;
                    }
                }
                else if (isWordFound == false && ocrWords[0].Contains("-"))
                {
                    string splittedCustomerNameWithoutHyphen = string.Concat(ocrWords[0].Split('-')[0], " ", ocrWords[0].Split('-')[1]);
                    if (DBName.Contains(splittedCustomerNameWithoutHyphen))
                    {
                        // isWordFound = true;
                        matchcount++;
                    }
                }

            }
            //if (isWordFound == true)
            //{
            //    IsWordMatchFromName = "Y";
            //}
            return matchcount;
        }
        public DataTable GetAliasDataFromMaster(string As400_ConnectionString, string State, string ZipCode, string AddressLine1, string City)
        {
            DataTable dtAlias = new DataTable();
            try
            {
                string queryStringShipper = "select * from ARP001 WHERE  CMSTAT = 'A' AND CMST ='" + State + "' AND CMZIP='" + ZipCode + "' AND CMCITY like '" + City + "%' AND CMADR1 like '" + AddressLine1 + "%'";
                dtAlias = GetDataTableByText(As400_ConnectionString, queryStringShipper);

            }
            catch (Exception ex)
            {
                dtAlias = null;
            }
            return dtAlias;
        }

        // added on 13-July-2021
        #region "Get ShipperKM details"
        private DataTable GetShipperKM(string spName, SqlParameter[] parameters = null)
        {
            DataTable dtShipperKM = new DataTable();
            try
            {
                string SQLConnectionString = string.Empty;
                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("SQLConn") == false)
                {
                    SqlConnection sqlConnection = new SqlConnection();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_SQLConn") == true)
                    {
                        SQLConnectionString = (string)iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_SQLConn"];
                    }
                    else
                    {
                        SQLConnectionString = ConfigurationManager.ConnectionStrings["ConnectionString_SQLConn"].ToString();
                    }
                    sqlConnection.ConnectionString = SQLConnectionString;

                    if (sqlConnection.State != ConnectionState.Open)
                    {
                        sqlConnection.Open();
                        iQPUBLIC.PublicComponents.htMyVariable.Add("SQLConn", sqlConnection);
                    }
                }

                SqlConnection connection = (SqlConnection)iQPUBLIC.PublicComponents.htMyVariable["SQLConn"];
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("SQLConn") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("SQLConn");
                    }
                    iQPUBLIC.PublicComponents.htMyVariable.Add("SQLConn", connection);
                }


                using (SqlCommand cmd = new SqlCommand(spName, connection))
                {
                    if (connection != null && !string.IsNullOrEmpty(spName))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters);
                        }

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dtShipperKM);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("GetSpecialInstructDetails: " + ex.Message);
            }
            return dtShipperKM;
        }

        #endregion
    }
}