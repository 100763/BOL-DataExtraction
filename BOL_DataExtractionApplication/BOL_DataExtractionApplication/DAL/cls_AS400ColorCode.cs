using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using BOLDocumentDataExtraction;
using BOLDocumentDataExtraction.CommonModule;

namespace BOLDocumentDataExtraction.DAL
{
    public static class cls_AS400ColorCode
    {
        /// <summary>
        /// Field list mapping with AS4000 color code fields
        /// </summary>
        /// <returns></returns>
        public static  Dictionary<string, string> _HeaderFieldsMappingToAS400_Back()
        {
            Dictionary<string, string> dicHeaderFieldsMapping = new Dictionary<string, string>();

            dicHeaderFieldsMapping.Add("BOL_No", "FHBL");
            dicHeaderFieldsMapping.Add("Shp_No", "FHSNUM");
            dicHeaderFieldsMapping.Add("PO_No", "FHPO");
            dicHeaderFieldsMapping.Add("Shp_Code", "FHSCD");
            dicHeaderFieldsMapping.Add("Shp_Name", "FHSNM");
            dicHeaderFieldsMapping.Add("Shp_Addr_Line", "FHSA1#FHSA2");
            dicHeaderFieldsMapping.Add("Shp_City", "FHSCT");
            dicHeaderFieldsMapping.Add("Shp_State", "FHSST");
            dicHeaderFieldsMapping.Add("Shp_ZipCode", "FHSZIP");
            dicHeaderFieldsMapping.Add("Con_Code", "FHCCD");
            dicHeaderFieldsMapping.Add("Con_Name", "FHCNM");
            dicHeaderFieldsMapping.Add("Con_Addr_Line", "FHCA1#FHCA2");
            dicHeaderFieldsMapping.Add("Con_City", "FHCCT");
            dicHeaderFieldsMapping.Add("Con_ZipCode", "FHCZIP");
            dicHeaderFieldsMapping.Add("Bill_To_Code", "FHBTC");
            dicHeaderFieldsMapping.Add("Bill_To_Name", "FHBNM");
            dicHeaderFieldsMapping.Add("Bill_To_Addr_Line", "FHBA1#FHBA2");
            dicHeaderFieldsMapping.Add("Bill_To_City", "FHBCT");
            dicHeaderFieldsMapping.Add("Bill_To_State", "FHBST");
            dicHeaderFieldsMapping.Add("Bill_To_ZipCode", "FHBZIP");
            dicHeaderFieldsMapping.Add("Terms", "FHTRM");
            dicHeaderFieldsMapping.Add("Total_Pieces", "FHTOTP");
            dicHeaderFieldsMapping.Add("Total_Weight", "FHSWGT");
            dicHeaderFieldsMapping.Add("Dates_RADF", "FHRADF");
            dicHeaderFieldsMapping.Add("Dates_MABD", "FHRADT");
            dicHeaderFieldsMapping.Add("FHCOD", "COD");

            return dicHeaderFieldsMapping;
        }

        public static Dictionary<string, string> _HeaderFieldsMappingToAS400()
        {
            Dictionary<string, string> dicHeaderFieldsMapping = new Dictionary<string, string>();

            dicHeaderFieldsMapping.Add("BOL_No", "FHBL");
            dicHeaderFieldsMapping.Add("Shp_No", "FHSNUM");
            dicHeaderFieldsMapping.Add("PO_No", "FHPO");
            dicHeaderFieldsMapping.Add("Shp_Code", "FHSCD");
            dicHeaderFieldsMapping.Add("Shp_Name", "FHSNM");
            dicHeaderFieldsMapping.Add("Shp_Addr_Line", "FHSA1#FHSA2");
            dicHeaderFieldsMapping.Add("Shp_City", "FHSCT");
            dicHeaderFieldsMapping.Add("Shp_State", "FHSST");
            dicHeaderFieldsMapping.Add("Shp_ZipCode", "FHSZIP");
            dicHeaderFieldsMapping.Add("Con_Code", "FHCCD");
            dicHeaderFieldsMapping.Add("Con_Name", "FHCNM");
            dicHeaderFieldsMapping.Add("Con_Addr_Line", "FHCA1#FHCA2");
            dicHeaderFieldsMapping.Add("Con_City", "FHCCT");
            dicHeaderFieldsMapping.Add("Con_State", "FHCST");
            dicHeaderFieldsMapping.Add("Con_ZipCode", "FHCZIP");
            dicHeaderFieldsMapping.Add("Bill_To_Code", "FHBTC");
            dicHeaderFieldsMapping.Add("Bill_To_Name", "FHBNM");
            dicHeaderFieldsMapping.Add("Bill_To_Addr_Line", "FHBA1#FHBA2");
            dicHeaderFieldsMapping.Add("Bill_To_City", "FHBCT");
            dicHeaderFieldsMapping.Add("Bill_To_State", "FHBST");
            dicHeaderFieldsMapping.Add("Bill_To_ZipCode", "FHBZIP");
            dicHeaderFieldsMapping.Add("Terms", "FHTRM");
            dicHeaderFieldsMapping.Add("Total_Pieces", "FHTOTP");
            dicHeaderFieldsMapping.Add("Total_Weight", "FHSWGT");
            dicHeaderFieldsMapping.Add("Dates_RADF", "FHRADF");
            dicHeaderFieldsMapping.Add("Dates_MABD", "FHRADT");
            dicHeaderFieldsMapping.Add("FHCOD", "COD");

            return dicHeaderFieldsMapping;
        }


        /// <summary>
        /// Method to insert color code data into AS400 FRPD010 table
        /// </summary>
        /// <param name="dtHeaderFldDtls"></param>
        /// <param name="ProNumber"></param>
        /// <param name="ConnectionString"></param>
        /// <returns></returns>
        public static bool InsertHeaderColorCodeIntoFRPD010(DataTable dtHeaderFldDtls, string ProNumber, string AS400DbConnectionString,string executionServiceName, long documentID)
        {
            Dictionary<string, string> dicHdFldMapping =  _HeaderFieldsMappingToAS400();
            System.Text.StringBuilder sbHeaderFieldXml = new System.Text.StringBuilder();
            DataTable dtSortedTable = new DataTable();
            dtSortedTable.Columns.Add("ProNumber");
            dtSortedTable.Columns.Add("FileName");
            dtSortedTable.Columns.Add("FieldName");
            dtSortedTable.Columns.Add("LineNumber");
            try
            {
                if (dtHeaderFldDtls != null && dtHeaderFldDtls.Rows.Count > 0)
                {
                    var result = from r in dtHeaderFldDtls.AsEnumerable()
                                 where r.Field<string>("FieldColor").ToString().ToLower().Trim().Contains("1")
                                 select r;
                    DataTable dtResult = result.CopyToDataTable();
                    if (dtResult != null)
                    {
                        for (int k = 0; k < dtResult.Rows.Count; k++)
                        {
                            var rows = dicHdFldMapping.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                                    == dtResult.Rows[k]["FieldName"].ToString().ToLower().Trim()
                                                                                    && dtResult.Rows[k]["FieldColor"].ToString().ToLower().Trim().Contains("1")).ToList();
                            if (rows.Count > 0)
                            {
                                if (rows[0].Value.ToString().Contains("#"))
                                {
                                    dtSortedTable.Rows.Add(
                                                   new string[] {
                                                ProNumber.ToString(),
                                                "FRP001",
                                                rows[0].Value.ToString().Split('#')[0],
                                                "0"
                                                   });

                                    dtSortedTable.Rows.Add(
                                                           new string[] {
                                                        ProNumber.ToString(),
                                                        "FRP001",
                                                        rows[0].Value.ToString().Split('#')[1],
                                                        "0"
                                                           });
                                }
                                else
                                {
                                    dtSortedTable.Rows.Add(new string[] {ProNumber.ToString(),"FRP001",rows[0].Value.ToString(),"0"
                                                       });
                                }
                            }
                        }
                        if(dtSortedTable!=null && dtSortedTable.Rows.Count>0)
                        {
                            InsertIntoFRPD010(documentID, dtSortedTable, AS400DbConnectionString);
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                BOLDocumentDataExtraction.DataExtraction.InsertErrorLog_As400(documentID, ErrorType.S, "cls_AS400ColorCode.InsertHeaderColorCodeIntoFRPD010: " + ex.Message, executionServiceName, "InsertHeaderColorCodeIntoFRPD010"); // DocumentID as 0
                BOLDocumentDataExtraction.DataExtraction.WriteLog(documentID + " cls_AS400ColorCode.InsertHeaderColorCodeIntoFRPD010 -> " + ex.Message.ToString(), LogType.E);

                //cls_CommonHelper.func_ProcessLog(AppConstant.ExceptionIn + " cls_AS400ColorCode.InsertHeaderIntoFRPD010 -> " + ex.Message.ToString());
                //InsertErrorLog(0, drsl_SAIABOL.Model.AppConstant.ErrorType.S, string.Concat("Login : ", Environment.UserName, " Error occurred - ", ex.Message.ToString()), "ClientUI", "cls_AS400ColorCode.InsertHeaderIntoFRPD010", Environment.MachineName);
                return false;
            }
            finally 
            {
                dtSortedTable.Dispose();
                dtHeaderFldDtls.Dispose();                 
            }
            return true;
        }
        public static int InsertIntoFRPD010_Back(long documentID, DataTable dataTable, string connectionString, OdbcParameter[] parameters = null)
        {
            int resultVal = 0;
            string queryString = "";
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    if (dataTable != null && dataTable.Rows.Count > 0)
                    {
                        if (connection.State == ConnectionState.Closed)
                        { connection.Open(); }
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            queryString = "insert into FRPD010(HXDPRO, HXDFIL, HXDFLD, HXDLIN) ";
                            queryString += " values('" + dataTable.Rows[i]["ProNumber"] + "', ";
                            queryString += "'" + dataTable.Rows[i]["FileName"] + "', ";
                            queryString += "'" + dataTable.Rows[i]["FieldName"] + "', ";
                            queryString += "'" + dataTable.Rows[i]["LineNumber"] + "' ";
                            queryString += ")";
                            using (OdbcCommand cmd = new OdbcCommand(queryString, connection))
                            {

                                if (connection != null && !string.IsNullOrEmpty(queryString))
                                {
                                    cmd.CommandType = CommandType.Text;
                                    if (parameters != null)
                                    {
                                        cmd.Parameters.AddRange(parameters);
                                    }
                                    resultVal = cmd.ExecuteNonQuery();
                                }

                            }
                        }
                    }
                    if (connection.State == ConnectionState.Open)
                    { connection.Close(); }
                }
                return resultVal;
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog(documentID + " cls_AS400ColorCode.InsertIntoFRPD010 -> " + ex.Message.ToString(), LogType.E);
                //cls_CommonHelper.func_ProcessLog(AppConstant.ExceptionIn + " DBUtilities.InsertOrUpdateBySP -> " + ex.Message.ToString());
                return resultVal;
            }
        }

        public static int InsertIntoFRPD010(long documentID, DataTable dataTable, string connectionString)
        {
            int resultVal = 0;
            string queryString = "";
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    if (dataTable != null && dataTable.Rows.Count > 0)
                    {
                        if (connection.State == ConnectionState.Closed)
                        { connection.Open(); }
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            queryString = "INSERT INTO FRPD010 (HXDPRO, HXDFIL, HXDFLD, HXDLIN) VALUES (?,?,?,?)";
                            if (connection != null && !string.IsNullOrEmpty(queryString))
                            {
                                using (OdbcCommand command = new OdbcCommand(queryString, connection))
                                {
                                    command.Parameters.Add("HXDPRO", OdbcType.VarChar).Value = dataTable.Rows[i]["ProNumber"];
                                    command.Parameters.Add("HXDFIL", OdbcType.VarChar).Value = dataTable.Rows[i]["FileName"];
                                    command.Parameters.Add("HXDFLD", OdbcType.VarChar).Value = dataTable.Rows[i]["FieldName"];
                                    command.Parameters.Add("HXDLIN", OdbcType.VarChar).Value = dataTable.Rows[i]["LineNumber"];
                                    resultVal = command.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                    if (connection.State == ConnectionState.Open)
                    { connection.Close(); }
                }
                return resultVal;
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog(documentID + " cls_AS400ColorCode.InsertIntoFRPD010 -> " + ex.Message.ToString(), LogType.E);
                // cls_CommonHelper.func_ProcessLog(AppConstant.ExceptionIn + "cls_AS400ColorCode.InsertOrUpdateBySP -> " + ex.Message.ToString());
                return resultVal;
            }
        }

        ///// <summary>
        ///// Method to insert error log
        ///// </summary>C:\DGSL_Project\drsl_SAIABOL\drsl_SAIABOL\drsl_SAIABOL\BAL\DocumentUtility.cs
        ///// <param name="documentID"></param>
        ///// <param name="errorType"></param>
        ///// <param name="errorDetails"></param>
        ///// <param name="moduleName"></param>
        ///// <param name="functionName"></param>
        //public static void InsertErrorLog(long documentID, drsl_SAIABOL.Model.AppConstant.ErrorType errorType, string errorDetails, string moduleName, string functionName, string ServerName)
        //{
        //    ErrorLogModel errorlogModel;
        //    try
        //    {
        //        errorlogModel = new ErrorLogModel
        //        {
        //            DocumentID = documentID,
        //            ErrorType = errorType,
        //            ErrorDetails = errorDetails,
        //            ModuleName = moduleName,
        //            FunctionName = functionName,
        //            ServerName = ServerName
        //        };

        //        if (!string.IsNullOrEmpty(drsl_SAIABOL.Model.AppConstant.ProductionConnString))
        //        {
        //            cls_CommonHelper.InsertErrorLog(errorlogModel, drsl_SAIABOL.Model.AppConstant.ProductionConnString);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        cls_CommonHelper.func_ProcessLog(AppConstant.ExceptionIn + " cls_AS400ColorCode.InsertErrorLog -> " + ex.Message.ToString());
        //    }
        //}


    }
}
