using System;
using System.Collections;
using System.Threading;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Data.Odbc;
using BOLDocumentDataExtraction.CommonModule;

namespace BOLDocumentDataExtraction.AS400ColorCoding

{
    public class cls_SaiaBol
    {
        static string str_FileName = string.Empty;
        static string Freight_Number = string.Empty;
        public static string str_SessionName;


        #region Live

        
        /// <summary>
        /// get Headerdata,LineItem data,Freight number from ClientUI,
        ///sort header and line item data in keyvalue pair
        ///read text from SAIA using readstr() function 
        ///filtered header and lineitem  data set to SAIA screen
        ///Set text for perticular field on SAIA using SetText() function 
        ///text has been set against row columns
        ///cls_ReturnStatus return function  class
        /// </summary>
        /// <param name="dtHeaderData"></param>
        /// <param name="dtLineitemdata"></param>
        /// <param name="FrightNumber"></param>
        /// <returns></returns>
        public static cls_ReturnStatus Func_GetDataFor_Freight(DataTable dt_HeaderData, DataTable dt_Lineitemdata, string FreightNumber, string SessionName,string AS400Con,string billingStatus)
        {

            #region Global Variables Declaration 

            DataSet ds = new DataSet();
            OdbcParameter[] parameters = null;
            DataTable dtHeaderData = new DataTable();
            DataTable dtLineitemdata = new DataTable();
            DataTable dtHeaderFilter = new DataTable();
            DataTable dtLineItemsFilter = new DataTable();
            cls_ReturnStatus _ReturnStatus = new cls_ReturnStatus();
            //cls_AppConstant.

            //int ThreadSleepTime = 0;
            string Connectionstr = string.Empty;
            //string str_SAIA_FreightNumber = string.Empty;
            string str_FreightNumber = "";

            List<KeyValuePair<String, String>> ListFilter = new List<KeyValuePair<String, String>>();
            List<KeyValuePair<string, List<string>>> LineItemsdictionary = new List<KeyValuePair<string, List<string>>>();
            #endregion

            BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
            BOLDocumentDataExtraction.DataExtraction.WriteLog("Func_GetDataFor_Freight Started", LogType.P);
            BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
           // BOLDocumentDataExtraction.DataExtraction.WriteLog("Session Name from Client UI--->" + SessionName);

            try
            {

                #region  Value Assign To Global Variables

                Connectionstr = AS400Con;//System.Configuration.ConfigurationManager.AppSettings["AS400DBConnString"].ToString();
                // ThreadSleepTime = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["ThreadSleepTime"].ToString());
                //cls_SaiaBol.str_SessionName = SessionName.Trim();
                #endregion




                #region Apply Rules On Fields 


                ds = cls_FieldRuleValidation.ApplyRulesOnFileds(dt_HeaderData, dt_Lineitemdata);
                if (ds.Tables.Count > 0)
                {
                    dtHeaderData = ds.Tables["Table1"];
                    if (dtHeaderData.Rows.Count > 0)
                    {
                        dtLineitemdata = ds.Tables["Table2"];
                    }
                }
                #endregion


                BOLDocumentDataExtraction.DataExtraction.WriteLog("Freight Bill Number Matched With AS400", LogType.P);
                BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);




                #region Filter & Set Data To AS400

                #region Header Filter 

                try
                {
                    //Filter datatable records having only true flag

                    BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Filter datatable records having only true flag for HeaderItem Fileds", LogType.P);
                    if (dtHeaderData.Rows.Count > 0)
                    {
                        var dtHeaderFiltervar = dtHeaderData.AsEnumerable()
                                       .Where(r => r.Field<string>("FieldColor").ToLower().ToString().Trim() == cls_AppConstant.FieldColorOne.ToLower().ToString().Trim())
                                       .ToList();
                        if (dtHeaderFiltervar.Any())
                        {
                            dtHeaderFilter = dtHeaderFiltervar.CopyToDataTable();
                            BOLDocumentDataExtraction.DataExtraction.WriteLog("Filter datatable records having only true flag Done for HeaderItem field and Count: " + dtHeaderFilter.Rows.Count, LogType.P);
                        }

                    }
                    //Adding datatable to List in key value pair
                    
                    if (dtHeaderFilter.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtHeaderFilter.Rows)
                        {
                            ListFilter.Add(new KeyValuePair<String, String>(dr[1].ToString(), Convert.ToString(dr[2])));
                        }
                    }
                    //StringBuilder sb = new StringBuilder();
                    //foreach (var element in ListFilter)
                    //{
                    //    sb.AppendLine(element.ToString());
                    //}
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Header fields filtered sucessfully", LogType.P);

                    ListFilter = cls_FieldRuleValidation.MatchAccountCode(ListFilter);
                }
                catch (Exception ex)
                {
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Excepetion in HeaderFilter--->" + ex, LogType.P);
                }
                #endregion

                #region Line Items Filter

                try
                {
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Filter datatable records having only true for LineItem Fileds",LogType.P);
                    if (dtLineitemdata.Rows.Count > 0)
                    {

                        var VardtLineItemsFilter = dtLineitemdata.AsEnumerable()
                                         .Where(r => r.Field<string>("FieldColor") == cls_AppConstant.FieldColorOne || r.Field<string>("FieldColor") == "True")
                                         .ToList();
                        if (VardtLineItemsFilter.Any())
                        {
                            dtLineItemsFilter = VardtLineItemsFilter.CopyToDataTable();
                            BOLDocumentDataExtraction.DataExtraction.WriteLog("Filter datatable records having only true flag Done for LineItem field and Count: " + dtLineItemsFilter.Rows.Count, LogType.P);
                        }
                    }
                    //Adding datatable to List in key value pair

                    if (dtLineItemsFilter.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtLineItemsFilter.Rows)
                        {
                            List<string> LineItemsValues = new List<string>();
                            LineItemsValues.Add(dr[1].ToString().Trim());
                            LineItemsValues.Add(dr[2].ToString());

                            LineItemsdictionary.Add(new KeyValuePair<String, List<string>>(dr[4].ToString(), LineItemsValues));
                        }
                    }
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Line Items filtered sucessfully", LogType.P);



                }
                catch (Exception ex)
                {
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in LineItems Filter--->" + ex,LogType.E);
                }
                BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------",LogType.P );
                #endregion


                Dictionary<string, string> HeaderItemFRP001 = new Dictionary<string, string>();

                #region Freight Number

                try
                {
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------",LogType.P );
                     str_FreightNumber = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Pro_No.ToLower().Trim())).Value;//String.Empty; //"10420322380";  

                    if (str_FreightNumber.Trim() == "")
                    {
                        str_FreightNumber = FreightNumber;
                    }
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Freight Number From UI --> " + str_FreightNumber, LogType.P);
                    //str_SAIA_FreightNumber = cls_Global.ReadStr(5, 11, 12, cls_SaiaBol.str_SessionName).Trim();
                    //BOLDocumentDataExtraction.DataExtraction.WriteLog("Freight Number From AS400 --> " + str_SAIA_FreightNumber);
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                }
                catch (Exception ex)
                {
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in read frieght number From Saia Screen section--->" + ex,LogType.E );
                }
                #endregion

                #region Variable Intialization

                string str_Shipper_Account_Code = String.Empty;//1
                string str_Shipper_Name = String.Empty;//2a
                string str_Shipper_Address = String.Empty;//2b
                string str_Shipper_Address2 = String.Empty;//2c
                string str_Shipper_City = String.Empty;//2d
                string str_Shipper_State = String.Empty;//2e
                string str_Shipper_Zip = String.Empty;//2f

                string str_Consignee_Account_Code = String.Empty;//3
                string str_Consignee_Name = String.Empty;//4a
                string str_Consignee_Address = String.Empty;//4b
                string str_Consignee_Address2 = String.Empty;//4c
                string str_Consignee_City = String.Empty;//4d
                string str_Consignee_State = String.Empty;//4e
                string str_Consignee_Zip = String.Empty;//4f

                string str_3pt_Bill_to_Account_Code = String.Empty;//5
                string str_3pt_Bill_to_Name = String.Empty;//6a
                string str_3pt_Bill_to_Address = String.Empty;//6b
                string str_3pt_Bill_to_Address2 = String.Empty;//6c
                string str_3pt_Bill_to_City = String.Empty;//6d
                string str_3pt_Bill_to_State = String.Empty;//6e
                string str_3pt_Bill_to_Zip = String.Empty;//6f


                string str_Terms = String.Empty;//7
                string str_Total_Pieces = String.Empty;//8
                string str_Total_Weight = String.Empty;//9
                string str_Bill_of_Lading_Number = String.Empty;//10
                string str_COD = String.Empty;//11
                string str_Delivery_Date_Code_Field = String.Empty;//12
                string str_Required_Arrival_Date1 = String.Empty;//13
                string str_Required_Arrival_Date2 = String.Empty;//13
                string str_Purchase_Order_Number = String.Empty;//14
                string str_Shipper_Number = String.Empty;//15
                string str_Declared_Value = String.Empty;//16
                string str_Inv_Value = String.Empty;//17
                string str_DLV = string.Empty;//18


                string str_LineItem_Pieces = String.Empty;//18
                string str_LineItem_Class_Code = String.Empty;//19
                string str_LineItem_Hazardous_Materials = String.Empty;//20
                string str_LineItem_Description = String.Empty;//21
                string str_LineItem_Weight = String.Empty;//22
                string str_LineItem_Pallet = String.Empty;//23
                #endregion

                cls_FieldRuleValidation.dtSortedTable = cls_AppConstant.FieldsMappingForTableFRPD010();

                DataTable dt_HeaderItemFRP001 = new DataTable();

                #region HeaderField Data Set on AS400
                
                string HeaderValue = string.Empty;
                BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------",LogType.P );
                #region Shipper
                try
                {
                    str_Shipper_Account_Code = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Shp_Code.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Shipper_Account_Code"]).Trim();
                    str_Shipper_Name = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Shp_Name.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Shipper_Name"]).Trim();

                    str_Shipper_Address = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Shp_Addr_Line.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Shipper_Address2"]).Trim();
                                                                                                                                                                       // str_Shipper_Address2 = ListFilter.FirstOrDefault(x => x.Key.Contains("Shp_Addr")).Value; //Convert.ToString(dtHeader.Rows[0]["Shipper_Address"]).Trim();
                    str_Shipper_City = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Shp_City.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Shipper_City"]).Trim();
                    str_Shipper_State = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Shp_State.ToLower().Trim())).Value; // Convert.ToString(dtHeader.Rows[0]["Shipper_State"]).Trim();
                    str_Shipper_Zip = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Shp_ZipCode.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Shipper_Zip"]).Trim();

                    List<string> _listShipper = cls_FieldRuleValidation.FindAndSetAddress(str_Shipper_Address);


                    if (!string.IsNullOrEmpty(str_Shipper_Account_Code))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("Shipper_Account_Code UI --> " + str_Shipper_Account_Code,LogType.P );
                      //  cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Shp_Code, str_FreightNumber, Connectionstr);
                        //Thread.Sleep(100);

                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Shp_Code, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Shipper_Account_Code = cls_FieldRuleValidation.TrimFieldValueLength(str_Shipper_Account_Code, cls_AppConstant.Shp_Con_Bill_Code_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Shipper_Account_Code);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": Shipper_Account_Code --> " + str_Shipper_Account_Code, LogType.P);
                        }
                    }

                    

                    if (!string.IsNullOrEmpty(str_Shipper_Name))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("Shipper_Name UI --> " + str_Shipper_Name, LogType.P);
                       // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Shp_Name, str_FreightNumber, Connectionstr);
                        //Thread.Sleep(100);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Shp_Name, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Shipper_Name = cls_FieldRuleValidation.TrimFieldValueLength(str_Shipper_Name, cls_AppConstant.AddressLength);
                            HeaderItemFRP001.Add(HeaderValue, str_Shipper_Name);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": Shipper_Name --> " + str_Shipper_Name, LogType.P);
                        }
                    }


                    if (!string.IsNullOrEmpty(str_Shipper_Address))
                    {
                        if (_listShipper.Count > 0)
                        {
                            if (_listShipper.Count > 1)
                            {
                                str_Shipper_Address = _listShipper[0]; str_Shipper_Address2 = _listShipper[1];
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("Shipper_Address UI --> " + str_Shipper_Address, LogType.P);
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("Shipper_Address2 UI --> " + str_Shipper_Address2, LogType.P);

                               // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Shp_Addr_Line, str_FreightNumber, Connectionstr);
                                //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Shp_Addr_Line1, str_FreightNumber, Connectionstr);

                                HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Shp_Addr_Line, str_FreightNumber);
                                if (!string.IsNullOrEmpty(HeaderValue))
                                {
                                    str_Shipper_Address = cls_FieldRuleValidation.TrimFieldValueLength(str_Shipper_Address, cls_AppConstant.AddressLength);
                                    HeaderItemFRP001.Add(HeaderValue, str_Shipper_Address);
                                    BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": Shipper_Address --> " + str_Shipper_Address, LogType.P);
                                }

                                HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Shp_Addr_Line1, str_FreightNumber);
                                if (!string.IsNullOrEmpty(HeaderValue))
                                {
                                    str_Shipper_Address2 = cls_FieldRuleValidation.TrimFieldValueLength(str_Shipper_Address2, cls_AppConstant.AddressLength);
                                    HeaderItemFRP001.Add(HeaderValue, str_Shipper_Address2);
                                    BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": Shipper_Address2 --> " + str_Shipper_Address2, LogType.P);
                                }
                            }
                            else
                            {
                                str_Shipper_Address = _listShipper[0];
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("Shipper_Address UI --> " + str_Shipper_Address, LogType.P);
                               // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Shp_Addr_Line, str_FreightNumber, Connectionstr);

                                HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Shp_Addr_Line, str_FreightNumber);
                                if (!string.IsNullOrEmpty(HeaderValue))
                                {
                                    
                                    str_Shipper_Address = cls_FieldRuleValidation.TrimFieldValueLength(str_Shipper_Address, cls_AppConstant.AddressLength);
                                    HeaderItemFRP001.Add(HeaderValue, str_Shipper_Address);
                                    BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": Shipper_Address --> " + str_Shipper_Address, LogType.P);
                                }
                            }
                        }
                       

                        //Thread.Sleep(100);
 
                    }

                    

                    if (!string.IsNullOrEmpty(str_Shipper_City))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("Shipper_City UI --> " + str_Shipper_City, LogType.P);
                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Shp_City, str_FreightNumber, Connectionstr);
                        //Thread.Sleep(100);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Shp_City, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Shipper_City = cls_FieldRuleValidation.TrimFieldValueLength(str_Shipper_City, cls_AppConstant.City_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Shipper_City);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": Shipper_City --> " + str_Shipper_City, LogType.P);
                        }
                    }

                   

                    if (!string.IsNullOrEmpty(str_Shipper_State))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("Shipper_State UI --> " + str_Shipper_State, LogType.P);
                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Shp_State, str_FreightNumber, Connectionstr);
                        //Thread.Sleep(100);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Shp_State, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Shipper_State = cls_FieldRuleValidation.TrimFieldValueLength(str_Shipper_State, cls_AppConstant.State_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Shipper_State);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": Shipper_State --> " + str_Shipper_State, LogType.P);
                        }
                    }

                   

                    if (!string.IsNullOrEmpty(str_Shipper_Zip))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("Shipper_Zip UI --> " + str_Shipper_Zip, LogType.P);
                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Shp_ZipCode, str_FreightNumber, Connectionstr);
                        //Thread.Sleep(100);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Shp_ZipCode, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Shipper_Zip = cls_FieldRuleValidation.TrimFieldValueLength(str_Shipper_Zip, cls_AppConstant.ZipCode_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Shipper_Zip);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": Shipper_Zip --> " + str_Shipper_Zip, LogType.P);
                        }
                    }

                    
                }
                catch (Exception ex)
                { 
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in Shipper SetText Section--->" + ex,LogType.E);
                }
                #endregion

                BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                #region Consignee

                try
                {
                    str_Consignee_Account_Code = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Con_Code.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Consignee_Account_Code"]).Trim();
                    str_Consignee_Name = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Con_Name.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["sConsignee_Name"]).Trim();
                                                                                                                                                                 //str_Consignee_Address2 = ListFilter.FirstOrDefault(x => x.Key.Contains("Con_Addr")).Value; // Convert.ToString(dtHeader.Rows[0]["Consignee_Address"]).Trim();
                    str_Consignee_Address = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Con_Addr_Line.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Consignee_Address2"]).Trim();
                    str_Consignee_City = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Con_City.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Consignee_City"]).Trim();
                    str_Consignee_State = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Con_State.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Consignee_State"]).Trim();
                    str_Consignee_Zip = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Con_ZipCode.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Consignee_Zip"]).Trim();

                    List<string> _listConsignee = cls_FieldRuleValidation.FindAndSetAddress(str_Consignee_Address);

                    //List<string> _listShipaddress = cls_FieldRuleValidation.SetTextOnNewLine(str_Shipper_Address, cls_AppConstant.AddressLength);



                    if (!string.IsNullOrEmpty(str_Consignee_Account_Code))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("  Consignee_Account_Code UI --> " + str_Consignee_Account_Code,LogType.P );
                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Con_Code, str_FreightNumber, Connectionstr);

                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Con_Code, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Consignee_Account_Code = cls_FieldRuleValidation.TrimFieldValueLength(str_Consignee_Account_Code, cls_AppConstant.Shp_Con_Bill_Code_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Consignee_Account_Code);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Consignee_Account_Code --> " + str_Consignee_Account_Code, LogType.P);
                        }
                    }

                    if (!string.IsNullOrEmpty(str_Consignee_Name))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog(" Consignee_Name UI --> " + str_Consignee_Name, LogType.P);
                    
                       // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Con_Name, str_FreightNumber, Connectionstr);
                        

                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Con_Name, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Consignee_Name = cls_FieldRuleValidation.TrimFieldValueLength(str_Consignee_Name, cls_AppConstant.AddressLength);
                            HeaderItemFRP001.Add(HeaderValue, str_Consignee_Name);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": Consignee_Name --> " + str_Consignee_Name, LogType.P);
                        }
                    }

                   

                    if (!string.IsNullOrEmpty(str_Consignee_Address))
                    {
                        if (_listConsignee.Count > 0)
                        {
                            if (_listConsignee.Count > 1)
                            {
                                str_Consignee_Address = _listConsignee[0]; str_Consignee_Address2 = _listConsignee[1];

                                BOLDocumentDataExtraction.DataExtraction.WriteLog("  Consignee_Address UI --> " + str_Consignee_Address, LogType.P);
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("  Consignee_Address2 UI --> " + str_Consignee_Address2, LogType.P);

                               // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Con_Addr_Line, str_FreightNumber, Connectionstr);
                                //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Con_Addr_Line1, str_FreightNumber, Connectionstr);

                                HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Con_Addr_Line, str_FreightNumber);
                                if (!string.IsNullOrEmpty(HeaderValue))
                                {
                                    str_Consignee_Address = cls_FieldRuleValidation.TrimFieldValueLength(str_Consignee_Address, cls_AppConstant.AddressLength);
                                    HeaderItemFRP001.Add(HeaderValue, str_Consignee_Address);
                                    BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Consignee_Address --> " + str_Consignee_Address, LogType.P);
                                }

                                HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Con_Addr_Line1, str_FreightNumber);
                                if (!string.IsNullOrEmpty(HeaderValue))
                                {
                                    str_Consignee_Address2 = cls_FieldRuleValidation.TrimFieldValueLength(str_Consignee_Address2, cls_AppConstant.AddressLength);
                                    HeaderItemFRP001.Add(HeaderValue, str_Consignee_Address2);
                                    BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Consignee_Address2 --> " + str_Consignee_Address2, LogType.P);
                                }
                            }
                            else
                            {
                                str_Consignee_Address = _listConsignee[0];
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("  Consignee_Address UI --> " + str_Consignee_Address, LogType.P);
                              //  cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Con_Addr_Line, str_FreightNumber, Connectionstr);
                                HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Con_Addr_Line, str_FreightNumber);
                                if (!string.IsNullOrEmpty(HeaderValue))
                                {
                                    str_Consignee_Address = cls_FieldRuleValidation.TrimFieldValueLength(str_Consignee_Address, cls_AppConstant.AddressLength);
                                    HeaderItemFRP001.Add(HeaderValue, str_Consignee_Address);
                                    BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Consignee_Address --> " + str_Consignee_Address, LogType.P);
                                }
                            }
                        }
                         
                       
                        
                    }

                   


                    if (!string.IsNullOrEmpty(str_Consignee_City))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("Consignee_City--->" + str_Consignee_City, LogType.P);
                       
                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Con_City, str_FreightNumber, Connectionstr);
                        // Thread.Sleep(100);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Con_City, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Consignee_City = cls_FieldRuleValidation.TrimFieldValueLength(str_Consignee_City, cls_AppConstant.City_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Consignee_City);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": Consignee_City --->" + str_Consignee_City, LogType.P);
                        }
                    }

                    if (!string.IsNullOrEmpty(str_Consignee_State))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("  Consignee_State UI --> " + str_Consignee_State, LogType.P);

                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Con_State, str_FreightNumber, Connectionstr);
                        //Thread.Sleep(100);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Con_State, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Consignee_State = cls_FieldRuleValidation.TrimFieldValueLength(str_Consignee_State, cls_AppConstant.State_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Consignee_State);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Consignee_State --> " + str_Consignee_State, LogType.P);
                        }
                    }
                    

                    if (!string.IsNullOrEmpty(str_Consignee_Zip))
                    {
                   
                        BOLDocumentDataExtraction.DataExtraction.WriteLog( "  Consignee_Zip UI --> " + str_Consignee_Zip, LogType.P);

                       // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Con_ZipCode, str_FreightNumber, Connectionstr);
                        //Thread.Sleep(100);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Con_ZipCode, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Consignee_Zip = cls_FieldRuleValidation.TrimFieldValueLength(str_Consignee_Zip, cls_AppConstant.ZipCode_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Consignee_Zip);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Consignee_Zip --> " + str_Consignee_Zip, LogType.P);
                        }
                    }

                   

                }
                catch (Exception ex)
                {
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in Consignee SetText section--->" + ex, LogType.E);
                }
                #endregion

                BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                #region 3Party

                try
                {
                    str_3pt_Bill_to_Account_Code = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Bill_To_Code.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["3pt_Bill_to_Account_Code"]).Trim();
                    str_3pt_Bill_to_Name = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Bill_To_Name.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["3pt_Bill_to_Name"]).Trim();
                                                                                                                                                                       // str_3pt_Bill_to_Address2 = ListFilter.FirstOrDefault(x => x.Key.Contains("Bill_To_Addr")).Value; //Convert.ToString(dtHeader.Rows[0]["3pt_Bill_to_Address"]).Trim();
                    str_3pt_Bill_to_Address = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Bill_To_Addr_Line.ToLower().Trim())).Value; // Convert.ToString(dtHeader.Rows[0]["3pt_Bill_to_Address2"]).Trim();
                    str_3pt_Bill_to_City = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Bill_To_City.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["3pt_Bill_to_City"]).Trim();
                    str_3pt_Bill_to_State = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Bill_To_State.ToLower().Trim())).Value; // Convert.ToString(dtHeader.Rows[0]["3pt_Bill_to_State"]).Trim();
                    str_3pt_Bill_to_Zip = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Bill_To_ZipCode.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["3pt_Bill_to_Zip"]).Trim();


                    List<string> _list3pt_Bill = cls_FieldRuleValidation.FindAndSetAddress(str_3pt_Bill_to_Address);

                    

                    if (!string.IsNullOrEmpty(str_3pt_Bill_to_Account_Code))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("  Tpt_Bill_to_Account_Code UI --> " + str_3pt_Bill_to_Account_Code, LogType.P);
                        
                       // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Bill_To_Code, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Bill_To_Code, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_3pt_Bill_to_Account_Code = cls_FieldRuleValidation.TrimFieldValueLength(str_3pt_Bill_to_Account_Code, cls_AppConstant.Shp_Con_Bill_Code_length);
                            HeaderItemFRP001.Add(HeaderValue, str_3pt_Bill_to_Account_Code);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Bill_to_Account_Code --> " + str_3pt_Bill_to_Account_Code, LogType.P);
                        }
                    }


                    if (!string.IsNullOrEmpty(str_3pt_Bill_to_Name))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("   Bill_to_Name UI --> " + str_3pt_Bill_to_Name, LogType.P);
                        
                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Bill_To_Name, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Bill_To_Name, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_3pt_Bill_to_Name = cls_FieldRuleValidation.TrimFieldValueLength(str_3pt_Bill_to_Name, cls_AppConstant.AddressLength );
                            HeaderItemFRP001.Add(HeaderValue, str_3pt_Bill_to_Name);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":   Bill_to_Name --> " + str_3pt_Bill_to_Name, LogType.P);
                        }
                    }


                    if (!string.IsNullOrEmpty(str_3pt_Bill_to_Address))
                    {

                        if (_list3pt_Bill.Count > 0)
                        {
                            if (_list3pt_Bill.Count > 1)
                            {
                                str_3pt_Bill_to_Address = _list3pt_Bill[0]; str_3pt_Bill_to_Address2 = _list3pt_Bill[1];

                                BOLDocumentDataExtraction.DataExtraction.WriteLog("  Bill_to_Address UI --> " + str_3pt_Bill_to_Address, LogType.P);
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("   Bill_to_Address2 UI --> " + str_3pt_Bill_to_Address2, LogType.P);

                               // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Bill_To_Addr_Line, str_FreightNumber, Connectionstr);
                                //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Bill_To_Addr_Line1, str_FreightNumber, Connectionstr);

                                HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Bill_To_Addr_Line, str_FreightNumber);
                                if (!string.IsNullOrEmpty(HeaderValue))
                                {
                                    str_3pt_Bill_to_Address = cls_FieldRuleValidation.TrimFieldValueLength(str_3pt_Bill_to_Address, cls_AppConstant.AddressLength);
                                    HeaderItemFRP001.Add(HeaderValue, str_3pt_Bill_to_Address);
                                    BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Bill_to_Address --> " + str_3pt_Bill_to_Address, LogType.P);
                                }

                                HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Bill_To_Addr_Line1, str_FreightNumber);
                                if (!string.IsNullOrEmpty(HeaderValue))
                                {
                                    str_3pt_Bill_to_Address2 = cls_FieldRuleValidation.TrimFieldValueLength(str_3pt_Bill_to_Address2, cls_AppConstant.AddressLength);
                                    HeaderItemFRP001.Add(HeaderValue, str_3pt_Bill_to_Address2);
                                    BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Bill_to_Address2 --> " + str_3pt_Bill_to_Address2, LogType.P);
                                }
                            }
                            else
                            {
                                str_3pt_Bill_to_Address = _list3pt_Bill[0];

                                BOLDocumentDataExtraction.DataExtraction.WriteLog("  Bill_to_Address UI --> " + str_3pt_Bill_to_Address, LogType.P);

                               // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Bill_To_Addr_Line, str_FreightNumber, Connectionstr);

                                HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Bill_To_Addr_Line, str_FreightNumber);
                                if (!string.IsNullOrEmpty(HeaderValue))
                                {
                                    str_3pt_Bill_to_Address = cls_FieldRuleValidation.TrimFieldValueLength(str_3pt_Bill_to_Address, cls_AppConstant.AddressLength);
                                    HeaderItemFRP001.Add(HeaderValue, str_3pt_Bill_to_Address);
                                    BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Bill_to_Address --> " + str_3pt_Bill_to_Address, LogType.P);
                                }
                            }
                        }
                        
                        
                        //Thread.Sleep(100);
                        
                        //Thread.Sleep(100);
                    }

                    

                    if (!string.IsNullOrEmpty(str_3pt_Bill_to_City))
                    {
                       
                        BOLDocumentDataExtraction.DataExtraction.WriteLog( "  Bill_to_City UI --> " + str_3pt_Bill_to_Name, LogType.P);

                       // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Bill_To_City, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Bill_To_City, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_3pt_Bill_to_City = cls_FieldRuleValidation.TrimFieldValueLength(str_3pt_Bill_to_City, cls_AppConstant.City_length);
                            HeaderItemFRP001.Add(HeaderValue, str_3pt_Bill_to_City);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Bill_to_City --> " + str_3pt_Bill_to_Name, LogType.P);
                        }
                    }

                   

                    if (!string.IsNullOrEmpty(str_3pt_Bill_to_State))
                    {
                      
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("   Bill_to_State UI --> " + str_3pt_Bill_to_State, LogType.P);

                      
                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Bill_To_State, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Bill_To_State, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_3pt_Bill_to_State = cls_FieldRuleValidation.TrimFieldValueLength(str_3pt_Bill_to_State, cls_AppConstant.State_length);
                            HeaderItemFRP001.Add(HeaderValue, str_3pt_Bill_to_State);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":   Bill_to_State --> " + str_3pt_Bill_to_State, LogType.P);
                        }
                    }

                   

                    if (!string.IsNullOrEmpty(str_3pt_Bill_to_Zip))
                    {
                         
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("   Bill_to_Zip UI --> " + str_3pt_Bill_to_Zip, LogType.P);

                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Bill_To_ZipCode, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Bill_To_ZipCode, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_3pt_Bill_to_Zip = cls_FieldRuleValidation.TrimFieldValueLength(str_3pt_Bill_to_Zip, cls_AppConstant.ZipCode_length);
                            HeaderItemFRP001.Add(HeaderValue, str_3pt_Bill_to_Zip);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":   Bill_to_Zip UI --> " + str_3pt_Bill_to_Zip, LogType.P);
                        }
                    }

                }
                catch (Exception ex)
                {
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in ThirdParty section--->" + ex.Message.ToString(), LogType.E );
                }
                #endregion
                BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                #region Others
                try
                {

                    str_Terms = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Terms.ToLower().Trim())).Value; // Convert.ToString(dtHeader.Rows[0]["Terms"]).Trim();
                    str_Total_Pieces = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Pieces.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Total_Pieces"]).Trim();
                    str_Bill_of_Lading_Number = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Bol_No.ToLower().Trim())).Value; // Convert.ToString(dtHeader.Rows[0]["Bill_of_Lading_Number"]).Trim();
                    str_COD = ListFilter.FirstOrDefault(x => x.Key.Trim().Contains(cls_HeaderStaticField.COD.Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["COD"]).Trim();
                    str_Required_Arrival_Date1 = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Dates_RADF.ToLower().Trim())).Value; // Convert.ToString(dtHeader.Rows[0]["Required_Arrival_Date1"]).Trim();
                    str_Required_Arrival_Date2 = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Dates_MABD.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Required_Arrival_Date2"]).Trim();
                    str_Purchase_Order_Number = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.PO_No.ToLower().Trim())).Value; // Convert.ToString(dtHeader.Rows[0]["Purchase_Order_Number"]).Trim();
                    str_Shipper_Number = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Shp_No.ToLower().Trim())).Value; //Convert.ToString(dtHeader.Rows[0]["Shipper_Number"]).Trim();
                    str_Delivery_Date_Code_Field = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Del_Dates_PDT.ToLower().Trim())).Value; // Convert.ToString(dtHeader.Rows[0]["Delivery_Date_Code_Field"]).Trim();
                    str_Total_Weight = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Total_Weight.ToLower().Trim())).Value;  // Convert.ToString(dtHeader.Rows[0]["Total_Weight"]).Trim();

                    str_Declared_Value = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Declared_Value.ToLower().Trim())).Value; ;// Convert.ToString(dtHeader.Rows[0]["Declared_Value"]).Trim();
                    str_Inv_Value = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.Inv_Value.ToLower().Trim())).Value; // Convert.ToString(dtHeader.Rows[0]["Full_Value_Coverage"]).Trim();
                    str_DLV = ListFilter.FirstOrDefault(x => x.Key.ToLower().Trim().Contains(cls_HeaderStaticField.dlv.ToLower().Trim())).Value;

                    if (!string.IsNullOrEmpty(str_Terms))
                    {
                        
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("  Terms UI --> " + str_Terms, LogType.P);

                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Terms, str_FreightNumber, Connectionstr);

                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Terms, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Terms = cls_FieldRuleValidation.TrimFieldValueLength(str_Terms, cls_AppConstant.Terms_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Terms);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Terms --> " + str_Terms, LogType.P);
                        }
                    }

                    

                    if (!string.IsNullOrEmpty(str_Total_Pieces))
                    {
                       
                        BOLDocumentDataExtraction.DataExtraction.WriteLog( "  Total_Pieces UI --> " + str_Total_Pieces, LogType.P);

                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Pieces, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Pieces, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Total_Pieces = cls_FieldRuleValidation.TrimFieldValueLength(str_Total_Pieces, cls_AppConstant.Total_Pieces_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Total_Pieces);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Total_Pieces --> " + str_Total_Pieces, LogType.P);
                        }
                    }
                   

                    if (!string.IsNullOrEmpty(str_Total_Weight))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog( "  Total_Weight UI --> " + str_Total_Weight, LogType.P);
                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Total_Weight, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Total_Weight, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Total_Weight = cls_FieldRuleValidation.TrimFieldValueLength(str_Total_Weight, cls_AppConstant.Total_Weight_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Total_Weight);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Total_Weight --> " + str_Total_Weight, LogType.P);
                        }
                    }
                    

                    if (!string.IsNullOrEmpty(str_Bill_of_Lading_Number))
                    {
                        //string Bill_of_Lading_Number_AS400 = cls_Global.ReadStr(13, 5, 20, cls_SaiaBol.str_SessionName).Trim(); 
                        BOLDocumentDataExtraction.DataExtraction.WriteLog( " Bill_of_Lading_Number UI --> " + str_Bill_of_Lading_Number, LogType.P);
                   
                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Bol_No, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Bol_No, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Bill_of_Lading_Number = cls_FieldRuleValidation.TrimFieldValueLength(str_Bill_of_Lading_Number, cls_AppConstant.PO_BL_SP_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Bill_of_Lading_Number);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": Bill_of_Lading_Number --> " + str_Bill_of_Lading_Number, LogType.P);
                        }
                    }


                    if (!string.IsNullOrEmpty(str_Purchase_Order_Number))
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("  Purchase_Order_Number UI --> " + str_Purchase_Order_Number, LogType.P);

                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.PO_No, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.PO_No, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Purchase_Order_Number = cls_FieldRuleValidation.TrimFieldValueLength(str_Purchase_Order_Number, cls_AppConstant.PO_BL_SP_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Purchase_Order_Number);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Purchase_Order_Number --> " + str_Purchase_Order_Number, LogType.P);
                        }
                    }

                    if (!string.IsNullOrEmpty(str_Shipper_Number))
                    {

                        BOLDocumentDataExtraction.DataExtraction.WriteLog("  Shipper_Number UI --> " + str_Shipper_Number, LogType.P);

                       // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Shp_No, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Shp_No, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Shipper_Number = cls_FieldRuleValidation.TrimFieldValueLength(str_Shipper_Number, cls_AppConstant.PO_BL_SP_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Shipper_Number);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Shipper_Number --> " + str_Shipper_Number, LogType.P);
                        }
                    }


                    if (!string.IsNullOrEmpty(str_COD))
                    {

                        BOLDocumentDataExtraction.DataExtraction.WriteLog(" COD UI --> " + str_COD, LogType.P);
                        
                       // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.COD, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.COD, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_COD = cls_FieldRuleValidation.TrimFieldValueLength(str_COD, cls_AppConstant.COD_length);
                            HeaderItemFRP001.Add(HeaderValue, str_COD);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ": COD --> " + str_COD, LogType.P);
                        }
                    }

                   

                    if (!string.IsNullOrEmpty(str_Delivery_Date_Code_Field))
                    {
                      
                        BOLDocumentDataExtraction.DataExtraction.WriteLog( "  Delivery_Date_Code_Field UI --> " + str_Delivery_Date_Code_Field, LogType.P);
                       
                       // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Del_Dates_PDT, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Del_Dates_PDT, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                           
                            HeaderItemFRP001.Add(HeaderValue, str_Delivery_Date_Code_Field);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Dates_RADF --> " + str_Required_Arrival_Date1, LogType.P);
                        }
                    }
                   

                    if (!string.IsNullOrEmpty(str_Required_Arrival_Date1))
                    {
                        
                        BOLDocumentDataExtraction.DataExtraction.WriteLog(  "  Required_Arrival_Date1 UI --> " + str_Required_Arrival_Date1, LogType.P);
                        
                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Dates_RADF, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Dates_RADF, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Required_Arrival_Date1 = cls_FieldRuleValidation.TrimFieldValueLength(str_Required_Arrival_Date1, cls_AppConstant.RADF_RADT_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Required_Arrival_Date1);

                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Dates_RADF --> " + str_Required_Arrival_Date1, LogType.P);
                        }
                    }

                    

                    if (!string.IsNullOrEmpty(str_Required_Arrival_Date2))
                    {
                       
                        BOLDocumentDataExtraction.DataExtraction.WriteLog( "  Required_Arrival_Date2 UI --> " + str_Required_Arrival_Date2, LogType.P);

                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Dates_MABD, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Dates_MABD, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            str_Required_Arrival_Date2 = cls_FieldRuleValidation.TrimFieldValueLength(str_Required_Arrival_Date2, cls_AppConstant.RADF_RADT_length);
                            HeaderItemFRP001.Add(HeaderValue, str_Required_Arrival_Date2);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Dates_MABD --> " + str_Required_Arrival_Date2, LogType.P);
                        }
                    }


                    if (!string.IsNullOrEmpty(str_DLV))
                    {

                        BOLDocumentDataExtraction.DataExtraction.WriteLog("  DLV_Value UI --> " + str_DLV, LogType.P);

                        //**************Change to table F2DLVA located in FRP001EX 11JAN22*************
                        //HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.dlv, str_FreightNumber);
                        //if (!string.IsNullOrEmpty(HeaderValue))
                        //{
                        //    str_DLV = cls_FieldRuleValidation.TrimFieldValueLength(str_DLV, cls_AppConstant.Dlv_length);
                        //    HeaderItemFRP001.Add(HeaderValue, str_DLV);
                        //    BOLDocumentDataExtraction.DataExtraction.WriteLog(HeaderValue + ":  Dlv --> " + str_DLV, LogType.P);
                        //}
                        //*******************************************************************************

                        str_DLV = cls_FieldRuleValidation.TrimFieldValueLength(str_DLV, cls_AppConstant.Dlv_length);


                    }


                    if (!string.IsNullOrEmpty(str_Declared_Value))
                    {
                      
                        BOLDocumentDataExtraction.DataExtraction.WriteLog ("  Declared_Value UI --> " + str_Declared_Value, LogType.P);
                        
                       // cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Declared_Value, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Declared_Value, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            HeaderItemFRP001.Add(HeaderValue, str_Declared_Value);
                        }
                    }
                   

                    if (!string.IsNullOrEmpty(str_Inv_Value))
                    {
                         
                        BOLDocumentDataExtraction.DataExtraction.WriteLog( "  Inv_Value UI --> " + str_Inv_Value, LogType.P);
                       
                        //cls_FieldRuleValidation.InsertHeaderColorCodeIntoFRPD010(cls_HeaderStaticField.Inv_Value, str_FreightNumber, Connectionstr);
                        HeaderValue = cls_FieldRuleValidation.HeaderItemMappingValue(cls_HeaderStaticField.Inv_Value, str_FreightNumber);
                        if (!string.IsNullOrEmpty(HeaderValue))
                        {
                            HeaderItemFRP001.Add(HeaderValue, str_Inv_Value);
                        }

                    }


                    
                    

                }
                catch (Exception ex) {BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in Others section--->" + ex,LogType.E); }


                #endregion
                BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);


                #endregion

                #region LineItems Data Set on AS400

                DataTable dt_LineItemFRP002 = new DataTable();
                //DataTable dt_LineItemFRP002 = new DataTable();
                string str_SAIACols = "Freight_Number|Code|Rate|PR|Amount|P|Description|Weight|PCS|H";
                string[] arr_SAIA_Cols = str_SAIACols.Split('|');
                foreach (string str_column in arr_SAIA_Cols)
                {
                    dt_LineItemFRP002.Columns.Add(str_column);
                }

               // dt_LineItemFRP002 = dt_SAIA.Clone();

                BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                #region ---READ DATA FROM CLIENT UI LINE ITEM---
                int LineItemsCount = dtLineItemsFilter.AsEnumerable().Select(r => r.Field<double>("RowNo")).Distinct().Count();

                var dt_rowcount = dtLineItemsFilter.AsEnumerable().Select(r => r.Field<double>("RowNo")).Distinct();

                ////  Entering Lines items

                int Max_LineNumber = 0;
                List<int> _lstlineitemNumber = new List<int>();

                try
                {
                    List<double> _ListRowCount = dt_rowcount.ToList();
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Read Data From Clien UI Table--> && Line Items Entered In SAIA", LogType.P);
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("For Freight Number: " + str_FreightNumber + ",Total Line Item Row Count: " + LineItemsCount, LogType.P);

                    
                    int LineNumber = cls_FieldRuleValidation.GetLineNumberFromFRP002(Connectionstr, str_FreightNumber);
                    Max_LineNumber = LineNumber;

                    for (int LineItem = 0; LineItem < LineItemsCount; LineItem++)
                    {


                        string Dt_RowNokey = Convert.ToString(_ListRowCount[LineItem]);
                        str_LineItem_Pieces = LineItemsdictionary.Where(d => d.Key == Convert.ToString(Dt_RowNokey) && d.Value[0].ToLower().Trim() == LineItemStaticField.Quantity.ToLower().Trim()).Select(d => d.Value[1]).FirstOrDefault();// Convert.ToString(dr["Quantity"]).Trim();
                        str_LineItem_Class_Code = LineItemsdictionary.Where(d => d.Key == Convert.ToString(Dt_RowNokey) && d.Value[0].ToLower().Trim() == LineItemStaticField.Class_Code.ToLower().Trim()).Select(d => d.Value[1]).FirstOrDefault(); //Convert.ToString(dr["Class_Code"]).Trim();
                        str_LineItem_Hazardous_Materials = LineItemsdictionary.Where(d => d.Key == Convert.ToString(Dt_RowNokey) && d.Value[0].ToLower().Trim() == LineItemStaticField.Hazmet_Code.ToLower().Trim()).Select(d => d.Value[1]).FirstOrDefault();// Convert.ToString(dr["HazMat_Code"]).Trim();
                        str_LineItem_Pallet = LineItemsdictionary.Where(d => d.Key == Convert.ToString(Dt_RowNokey) && d.Value[0].ToLower().Trim() == LineItemStaticField.Pallet_code.ToLower().Trim()).Select(d => d.Value[1]).FirstOrDefault();// Convert.ToString(dr["Pallet"]).Trim();
                        str_LineItem_Description = LineItemsdictionary.Where(d => d.Key == Convert.ToString(Dt_RowNokey) && d.Value[0].ToLower().Trim() == LineItemStaticField.Description.ToLower().Trim()).Select(d => d.Value[1]).FirstOrDefault();// Convert.ToString(dr["Description"]).Trim();
                        str_LineItem_Weight = LineItemsdictionary.Where(d => d.Key == Convert.ToString(Dt_RowNokey) && d.Value[0].ToLower().Trim() == LineItemStaticField.Weight.ToLower().Trim()).Select(d => d.Value[1]).FirstOrDefault();// Convert.ToString(dr["Weight"]).Trim();


                        BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("For Freight Number: " + str_FreightNumber + ",Line Item Row: " + LineItem + "---> " +
                                                    "Quantity: " + str_LineItem_Pieces + "" +
                                                    ",Pallet_Code: " + str_LineItem_Pallet + "" +
                                                    ",Clas_Code: " + str_LineItem_Class_Code + "" +
                                                    ",Description: " + str_LineItem_Description + "" +
                                                    ",Weight: " + str_LineItem_Weight + "" +
                                                    ",Rate:" + "" + ",PR: " + "" + ",Amount: " + "" + "" +
                                                    ",Hazmat_Code: " + str_LineItem_Hazardous_Materials, LogType.P);





                        LineNumber = LineNumber + 1;
                        _lstlineitemNumber.Add(LineNumber);

                        DataRow dr_SAIA = dt_LineItemFRP002.NewRow();


                        dr_SAIA["Freight_Number"] = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_FreightNumber);

                        //dr_SAIA["Code"] = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Class_Code);
                        string str_SAIA_Code = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Class_Code);

                        if (!string.IsNullOrEmpty(str_SAIA_Code))
                        {
                            //str_SAIA_Code = cls_FieldRuleValidation.Class_CodeValidate(str_SAIA_Code);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog("Class Code Value--->" + str_SAIA_Code, LogType.P);
                            cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(LineItemStaticField.Class_Code, str_FreightNumber, Connectionstr, LineNumber);
                        }
                        dr_SAIA["Code"] = str_SAIA_Code;


                        dr_SAIA["Rate"] = "";
                        dr_SAIA["PR"] = "";
                        dr_SAIA["Amount"] = "";

                        //****************Rate, PR & Amount value empty always.
                        //string str_SAIA_Rate = Convert.ToString(dr["Rate"]).Trim().Replace(".", "");
                        //string str_SAIA_PR = Convert.ToString(dr["PR"]).Trim();
                        //string str_SAIA_Amt = Convert.ToString(dr["Amount"]).Trim().Replace(".", "");

                        //if (!string.IsNullOrEmpty(str_SAIA_Rate))
                        //{
                        //    cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(null, str_FreightNumber, Connectionstr, LineNumber);
                        //}

                        //if (!string.IsNullOrEmpty(str_SAIA_PR))
                        //{
                        //    cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(null, str_FreightNumber, Connectionstr, LineNumber);
                        //}
                        //if (!string.IsNullOrEmpty(str_SAIA_Amt))
                        //{
                        //    cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(null, str_FreightNumber, Connectionstr, LineNumber);
                        //}
                        //***************************************************************

                        //dr_SAIA["P"] = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Pallet);
                        string str_SAIA_P = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Pallet);

                        if (!string.IsNullOrEmpty(str_SAIA_P))
                        {
                            cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(LineItemStaticField.Pallet_code, str_FreightNumber, Connectionstr, LineNumber);
                        }
                        dr_SAIA["P"] = str_SAIA_P;

                        string str_SAIA_Desc = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Description);
                        //dr_SAIA["Description"] = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Description);

                        if (!string.IsNullOrEmpty(str_SAIA_Desc))
                        {
                            cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(LineItemStaticField.Description, str_FreightNumber, Connectionstr, LineNumber);
                        }

                        dr_SAIA["Description"] = str_SAIA_Desc;

                        //dr_SAIA["Weight"] = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Weight).Replace(",", " ");
                        string str_SAIA_Wt = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Weight).Replace(",", " ");
                        if (!string.IsNullOrEmpty(str_SAIA_Wt))
                        {
                            cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(LineItemStaticField.Weight, str_FreightNumber, Connectionstr, LineNumber);
                        }
                        dr_SAIA["Weight"] = str_SAIA_Wt;


                        //dr_SAIA["PCS"] = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Pieces);
                        string str_SAIA_PCS = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Pieces);

                        if (!string.IsNullOrEmpty(str_SAIA_PCS))
                        {
                            string[] _qunatityArray = null;
                            if (!string.IsNullOrEmpty(str_SAIA_PCS))
                            {
                                if (str_SAIA_PCS.Contains(cls_AppConstant.DotSign))
                                {
                                    _qunatityArray = str_SAIA_PCS.Split('.');
                                    str_SAIA_PCS = _qunatityArray[0];

                                }
                            }
                            //dr["PCS"] = Convert.ToString(str_SAIA_PCS);
                            BOLDocumentDataExtraction.DataExtraction.WriteLog("Quantity Value--->" + str_SAIA_PCS, LogType.P);
                            cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(LineItemStaticField.Quantity, str_FreightNumber, Connectionstr, LineNumber);
                        }
                        dr_SAIA["PCS"] = str_SAIA_PCS;

                        //dr_SAIA["H"] = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Hazardous_Materials);
                        string str_SAIA_H = cls_FieldRuleValidation.CheckIsnullOrEmpty(str_LineItem_Hazardous_Materials);
                        if (!string.IsNullOrEmpty(str_SAIA_H))
                        {
                            //rowFinal["H"] = Convert.ToString(str_SAIA_H);
                            cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(null, str_FreightNumber, Connectionstr, LineNumber);
                        }
                        dr_SAIA["H"] = str_SAIA_H;


                        dt_LineItemFRP002.Rows.Add(dr_SAIA);
                        dt_LineItemFRP002.AcceptChanges();


                        BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("For Freight Number: " + str_FreightNumber + ",Line Item Row: -->" + LineNumber + "" +
                                                        "Quantity: " + str_SAIA_PCS + "" +
                                                        ",Pallet_Code: " + str_SAIA_P + "" +
                                                        ",Clas_Code: " + str_SAIA_Code + "" +
                                                        ",Description: " + str_SAIA_Desc + "" +
                                                        ",Weight: " + str_SAIA_Wt + "" +
                                                        ",Rate: " + "" + ",PR: " + "" + ",Amount: " + "" + "" +
                                                        ",Hazmat_Code: " + str_SAIA_H, LogType.P);

                    }
                }
                catch (Exception ex)
                {

                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in LineItems SetText Forloop--->" + ex, LogType.E);
                }
                #endregion
                BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);

                #region ---ENTER DATA IN SAIA---

               // //Thread.Sleep(100);
               //// cls_Global.Reset(); //Rajesh
               // BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
               // BOLDocumentDataExtraction.DataExtraction.WriteLog("Line Items Entered In SAIA", LogType.P);
               // BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
               // List<int> _lstlineitemNumber = new List<int>();
               // int LineNumber = cls_FieldRuleValidation.GetLineNumberFromFRP002(Connectionstr, str_FreightNumber);
               // int Max_LineNumber = LineNumber;
               // foreach (DataRow dr in dt_LineItemFRP002.Rows)
               // {
               //     //string str_SAIACols = "Freight_Number|Code|Rate|PR|Amount|P|Description|Weight|PCS|H";
               //     string str_SAIA_Code = Convert.ToString(dr["Code"]).Trim();

               //     string str_SAIA_Rate = Convert.ToString(dr["Rate"]).Trim().Replace(".", "");
               //     string str_SAIA_PR = Convert.ToString(dr["PR"]).Trim();
               //     string str_SAIA_Amt = Convert.ToString(dr["Amount"]).Trim().Replace(".", "");

               //     string str_SAIA_P = Convert.ToString(dr["P"]).Trim();
               //     string str_SAIA_Desc = Convert.ToString(dr["Description"]).Trim();
               //     string str_SAIA_Wt = Convert.ToString(dr["Weight"]).Trim();
               //     string str_SAIA_PCS = Convert.ToString(dr["PCS"]).Trim();
               //     string str_SAIA_H = Convert.ToString(dr["H"]).Trim();

               //     //DataRow rowFinal = dt_LineItemFRP002.NewRow();

               //     LineNumber = LineNumber + 1;
               //     _lstlineitemNumber.Add(LineNumber);

               //     BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
               //     BOLDocumentDataExtraction.DataExtraction.WriteLog("For Freight Number: " + str_FreightNumber + ",Line Item Row: -->" + LineNumber + "" +
               //                                     "Quantity: " + str_SAIA_PCS + "" +
               //                                     ",Pallet_Code: " + str_SAIA_P + "" +
               //                                     ",Clas_Code: " + str_SAIA_Code + "" +
               //                                     ",Description: " + str_SAIA_Desc + "" +
               //                                     ",Weight: " + str_SAIA_Wt + "" +
               //                                     ",Rate:" + str_SAIA_Rate + ",PR: " + str_SAIA_PR + ",Amount: " + str_SAIA_Amt + "" +
               //                                     ",Hazmat_Code: " + str_SAIA_H, LogType.P);


               //     #region ColorCode ClientUI Values

               //     if (!string.IsNullOrEmpty(str_SAIA_Code))
               //     {

               //         str_SAIA_Code = cls_FieldRuleValidation.Class_CodeValidate(str_SAIA_Code);
               //         dr["Code"] = Convert.ToString(str_SAIA_Code);
               //         //rowFinal["Code"] = Convert.ToString(str_SAIA_Code);
               //         BOLDocumentDataExtraction.DataExtraction.WriteLog("Class Code Value--->" + str_SAIA_Code, LogType.P);
               //         cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(LineItemStaticField.Class_Code, str_FreightNumber, Connectionstr, LineNumber);

               //     }
                   

               //     if (!string.IsNullOrEmpty(str_SAIA_Rate))
               //     {
               //         //rowFinal["Rate"] = Convert.ToString(str_SAIA_Rate);
               //         cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(null, str_FreightNumber, Connectionstr, LineNumber);
               //     }
                    

               //     if (!string.IsNullOrEmpty(str_SAIA_PR))
               //     {
               //        // rowFinal["PR"] = Convert.ToString(str_SAIA_PR);
               //         cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(null, str_FreightNumber, Connectionstr, LineNumber);
               //     }
               //     if (!string.IsNullOrEmpty(str_SAIA_Amt))
               //     {
               //         //rowFinal["Amount"] = Convert.ToString(str_SAIA_Amt);
               //         cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(null, str_FreightNumber, Connectionstr, LineNumber);
               //     }

               //     if (!string.IsNullOrEmpty(str_SAIA_P))
               //     {
               //         //rowFinal["P"] = Convert.ToString(str_SAIA_P);
               //         cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(LineItemStaticField.Pallet_code, str_FreightNumber, Connectionstr, LineNumber);
               //     }
               //     if (!string.IsNullOrEmpty(str_SAIA_Desc))
               //     {
               //         //rowFinal["Description"] = Convert.ToString(str_SAIA_Desc);
               //         cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(LineItemStaticField.Description, str_FreightNumber, Connectionstr, LineNumber);
               //     }
               //     if (!string.IsNullOrEmpty(str_SAIA_Wt))
               //     {
               //        // rowFinal["Weight"] = Convert.ToString(str_SAIA_Wt);
               //         cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(LineItemStaticField.Weight, str_FreightNumber, Connectionstr, LineNumber);
               //     }

               //     if (!string.IsNullOrEmpty(str_SAIA_PCS))
               //     {
               //         string[] _qunatityArray = null;
               //         if (!string.IsNullOrEmpty(str_SAIA_PCS))
               //         {
               //             if (str_SAIA_PCS.Contains(cls_AppConstant.DotSign))
               //             {
               //                 _qunatityArray = str_SAIA_PCS.Split('.');
               //                 str_SAIA_PCS = _qunatityArray[0];

               //             }
               //         }
               //         dr["PCS"] = Convert.ToString(str_SAIA_PCS);
               //        // rowFinal["PCS"] = Convert.ToString(str_SAIA_PCS);
               //         BOLDocumentDataExtraction.DataExtraction.WriteLog("Quantity Value--->" + str_SAIA_PCS, LogType.P);
               //         //TruBotMainframe.Mainframe.SetCursorPosition(j, int_SAIA_PCSColPOS, cls_SaiaBol.str_SessionName);
               //         //TruBotMainframe.Mainframe.SetText(str_SAIA_PCS, j, int_SAIA_PCSColPOS, cls_SaiaBol.str_SessionName);
               //         cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(LineItemStaticField.Quantity, str_FreightNumber, Connectionstr, LineNumber);

               //     }
               //     if (!string.IsNullOrEmpty(str_SAIA_H))
               //     {
               //         //rowFinal["H"] = Convert.ToString(str_SAIA_H);
               //         cls_FieldRuleValidation.InsertLineColorCodeIntoFRPD010(null, str_FreightNumber, Connectionstr, LineNumber);
               //     }

               //     dt_LineItemFRP002.AcceptChanges();
               //     //rowFinal = dr;
               //     //dt_LineItemFRP002.Rows.Add(rowFinal);
               //     //dt_LineItemFRP002.Rows.Add(dr);

               //     //try
               //     //{
               //     //    // OdbcParameter[] parameters = null;
               //     //    cls_FieldRuleValidation.InsertIntoFRP002(dr, LineNumber, Connectionstr, Freight_Number);
               //     //}
               //     //catch (Exception ex)
               //     //{
               //     //    cls_Global.func_ErrorLog("InsertIntoFRP002--->" + ex);
               //     //}


               //     #endregion

               // }

                #endregion

                #endregion

                #endregion



                if (cls_FieldRuleValidation.dtSortedTable != null && cls_FieldRuleValidation.dtSortedTable.Rows.Count > 0)
                {
                    _ReturnStatus = cls_FieldRuleValidation.UpdateIntoFRPD010_UpdateIntoFRP001_InsertIntoFRP002(HeaderItemFRP001, dt_LineItemFRP002, Max_LineNumber, _lstlineitemNumber, Connectionstr, str_FreightNumber, billingStatus, str_DLV, parameters);

                    if (cls_FieldRuleValidation.dtSortedTable != null && cls_FieldRuleValidation.dtSortedTable.Rows.Count > 0) { cls_FieldRuleValidation.dtSortedTable.Clear(); }

                    if (HeaderItemFRP001 != null && HeaderItemFRP001.Count > 0) { HeaderItemFRP001.Clear(); }

                    if (dt_LineItemFRP002 != null && dt_LineItemFRP002.Rows.Count > 0) { dt_LineItemFRP002.Clear(); }
                }
                else
                {

                    _ReturnStatus.Message = "Transfer Success";
                    _ReturnStatus.Status = "S";
                    _ReturnStatus.IsStatus = true;
                }



                BOLDocumentDataExtraction.DataExtraction.WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);

                BOLDocumentDataExtraction.DataExtraction.WriteLog("---------------------------------End of func_GetDataFor_Freight()---------------------------------", LogType.P);


                return _ReturnStatus;

            }
            catch (Exception ex)
            {

                _ReturnStatus.Message = cls_MessageCode.EX014;
                _ReturnStatus.Status = "F";
                _ReturnStatus.IsStatus = false;
                _ReturnStatus.Value = ex.Message;
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in func_GetDataFor_Freight()--->" + ex, LogType.E);
                return _ReturnStatus;
            }
        }

        public static void CheckSessionNameFromClient()
        {
            try
            {
                str_SessionName = System.Configuration.ConfigurationManager.AppSettings["SessionName"].ToString();
            }
            catch (Exception)
            {
            }
        }

        public static bool CheckIsSessionScreenFromAS()
        {
            try
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Session Name from Client UI--->" + cls_SaiaBol.str_SessionName,LogType.P );
                string str_ship = "";//cls_Global.ReadStr(6, 2, 4, cls_SaiaBol.str_SessionName).Trim();

                if (str_ship != "Ship")
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("CheckIsSessionFromAS--->" + ex,LogType.E);
                return false;
            }
            return true;

        }
        #endregion


      
        public static DataTable ReadExcel(string str_FilePath, string str_SheetName = null)
        {
            string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + str_FilePath + ";Extended Properties='Excel 8.0;HDR={1};IMEX=1'";
            string query = string.Empty;


            OleDbConnection conn = new OleDbConnection(connStr);
            OleDbCommand comm = new OleDbCommand();
            DataTable dtselect = new DataTable();
            if (conn.State == ConnectionState.Closed)
                conn.Open();
            if (!String.IsNullOrEmpty(str_SheetName))
            {
                query = "select * from [" + str_SheetName + "$]";
            }
            else
            {
                var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                query = "select * from [" + sheets.Rows[0]["TABLE_NAME"].ToString() + "] where [Flag] is null";
            }
            comm.Connection = conn;
            comm.CommandText = query;
            comm.CommandType = CommandType.Text;
            //comm.Parameters.Add(new OleDbParameter { Value = str_Flag });
            //comm.Parameters.Add(new OleDbParameter { Value = str_Remark });

            try
            {
                //comm.ExecuteNonQuery();
                OleDbDataAdapter ole = new OleDbDataAdapter(comm);
                ole.Fill(dtselect);
            }
            catch (System.Exception ex)
            {
                //cls_Global.func_ErrorLog("Exception in  Readexcel executing query: " + query + " is :" + Convert.ToString(ex));
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in Readexcel executing query: " + query + " is :" + ex.Message, LogType.E);
                //clsGlobal.func_ShowMessageBox("Exception in executing query: " + query + " is :" + ex.Message);
            }
            conn.Close();
            return dtselect;
        }

        public static void UpdateFreightExcel(string str_FilePath, string str_FreightNo)
        {
            string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + str_FilePath + ";Extended Properties=Excel 12.0 Xml";
            string query = string.Empty;

            OleDbConnection conn = new OleDbConnection(connStr);
            OleDbCommand comm = new OleDbCommand();

            if (conn.State == ConnectionState.Closed)
                conn.Open();
            var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            query = "UPDATE [" + sheets.Rows[0]["TABLE_NAME"].ToString() + "] SET [Flag]=? where [Freight Number]=?";

            comm.Connection = conn;
            comm.CommandText = query;
            comm.CommandType = CommandType.Text;
            comm.Parameters.Add(new OleDbParameter { Value = 1 });
            //comm.Parameters.Add(new OleDbParameter { Value = str_Remark });
            comm.Parameters.Add(new OleDbParameter { Value = str_FreightNo });
            try
            {
                comm.ExecuteNonQuery();
            }
            catch (System.Exception ex)
            {
               // BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in UpdateFreightExcel executing query: " + query + " is :" + Convert.ToString(ex),LogType.E );
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in UpdateFreightExcel executing query: " + query + " is :" + ex.Message, LogType.E);
                //clsGlobal.func_ShowMessageBox("Exception in executing query: " + query + " is :" + ex.Message);
            }
            conn.Close();
        }

        /// <summary>
        /// Kill background process in task manager
        /// </summary>
        public static void KillSpecificProcess()
        {
            try
            {
                var processes = from p in System.Diagnostics.Process.GetProcessesByName("PCSPCOC")
                                select p;

                foreach (var process in processes)
                {
                    //if (process.MainWindowTitle!=null && process.MainWindowTitle!=string.Empty ) //&& process.MainWindowTitle.ToLower().Contains("excel"))
                    try
                    {
                        process.Kill();
                        process.Close();
                    }
                    catch (System.Exception)
                    {


                    }

                }
            }
            catch (System.Exception ex)
            {
                //cls_Global.func_ErrorLog("KillSpecificProcess" + Convert.ToString(ex));
                BOLDocumentDataExtraction.DataExtraction.WriteLog("KillSpecificProcess:" + ex.Message,LogType.E);

            }
        }
        
    }
}
