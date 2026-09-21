using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using iQDataProvider;
using iQPUBLIC;

using BOLDocumentDataExtraction.DAL;
//using BOLDocumentDataExtraction.BLL;
using BOLDocumentDataExtraction.CommonModule;
using BOLDocumentDataExtraction.AS400ColorCoding;
using System.Collections;

namespace BOLDocumentDataExtraction
{
    public class DataExtraction
    {
        #region "Variable Decl"

        static string cryptographyKey;
        static string processLogfileName;
        static string processLogFilePath;

        static StreamWriter processLogWriter;
        static string executionServerName;
        static string executionServiceName;
        static List<AppSetting> destinationServerDetailsList;
        string Cons_Code_Misc;

        static string billingStatus;

        ClsCNC objCNC;
        DataTable dtfieldInfo;


        SAIABOL01IQ.clsF3 objKM2F3;
        SAIABOL01IQ.clsF5 objKM2F5;
        SAIABOL01IQ.clsF7 objKM2F7;
        SAIABOL01IQ.clsF27 objKM2F27;

        SAIABOL01IQ.clsF6 objKM2F6;
        SAIABOL01IQ.clsF23 objKM2F23;
        SAIABOL01IQ.clsF24 objKM2F24;

        RetStructF3 retStructF3Value;
        RetStructF27 retStructF27Value;

        StringBuilder objExtractedData;

       // DataTable dt_HeaderFileddata;
        //DataTable dt_lineitemdata;

        DataTable dt_SortedHeaderFielddata;
        DataTable dt_SortedLineitemdata;

        int SrNo;
        bool Shp_ChangeGreenToRed;
        bool Con_ChangeGreenToRed;
        bool BillTo_ChangeGreenToRed;

        string Pro_No;
        string Shp_Code;
        string Shp_Name;
        string Shp_Addr_Line;
        string Shp_City;
        string Shp_State;
        string Shp_ZipCode;

        string Con_Code;
        string Con_Name;
        string Con_Addr_Line;
        string Con_City;
        string Con_State;
        string Con_ZipCode;
        string Con_TelNo;

        string Bill_To_Code;
        string Bill_To_Name;
        string Bill_To_Addr_Line;
        string Bill_To_City;
        string Bill_To_State;
        string Bill_To_ZipCode;

        string Shp_Name_OBL;
        string Shp_Addr_Line_OBL;
        string Shp_City_OBL;
        string Shp_State_OBL;
        string Shp_ZipCode_OBL;

        string Shp_Code_PKP;
        string ShipperDetails_PKP;

        int PageCount = 0;

        Dictionary<string, string> ShipperList;
        Dictionary<string, clsCnCWord> ConsigneeList;// Added on 12aug21

        ArrayList arValidClasscode;


        #endregion

        public void StaticVariable_Initialize()
        {
           

            cryptographyKey = "MACKCP110216";
            processLogfileName = "process";
            processLogFilePath = "";
            executionServiceName = "BOLDataExtractionService";

            executionServerName = string.Empty;
            destinationServerDetailsList = new List<AppSetting>();


            dtfieldInfo = null;


            DBUtilities.appConnectionString = string.Empty;
            DBUtilities.prodConnectionString = string.Empty;
            DBUtilities.DB2ConnectionString = string.Empty;

            Cons_Code_Misc = string.Empty;

            //dt_HeaderFileddata = new DataTable();
            //dt_lineitemdata = new DataTable();

            dt_SortedHeaderFielddata = new DataTable();
            dt_SortedLineitemdata = new DataTable();


            dt_SortedHeaderFielddata.Columns.Add("SrNo");
            dt_SortedHeaderFielddata.Columns.Add("FieldName");
            dt_SortedHeaderFielddata.Columns.Add("FieldValue");
            dt_SortedHeaderFielddata.Columns.Add("FieldColor");

            dt_SortedLineitemdata.Columns.Add("SrNo");
            dt_SortedLineitemdata.Columns.Add("FieldName");
            dt_SortedLineitemdata.Columns.Add("FieldValue");
            dt_SortedLineitemdata.Columns.Add("FieldColor");
            dt_SortedLineitemdata.Columns.Add("RowNo", typeof(double));

            Shp_ChangeGreenToRed = false;
            Con_ChangeGreenToRed = false;
            BillTo_ChangeGreenToRed = false;

            Pro_No = "Pro_No";

            Shp_Code = "Shp_Code";
            Shp_Name = "Shp_Name";
            Shp_Addr_Line = "Shp_Addr_Line";
            Shp_City = "Shp_City";
            Shp_State = "Shp_State";
            Shp_ZipCode = "Shp_ZipCode";

            Con_Code = "Con_Code";
            Con_Name = "Con_Name";
            Con_Addr_Line = "Con_Addr_Line";
            Con_City = "Con_City";
            Con_State = "Con_State";
            Con_ZipCode = "Con_ZipCode";
            Con_TelNo = "Con_TelNo";

            Bill_To_Code = "Bill_To_Code";
            Bill_To_Name = "Bill_To_Name";
            Bill_To_Addr_Line = "Bill_To_Addr_Line";
            Bill_To_City = "Bill_To_City";
            Bill_To_State = "Bill_To_State";
            Bill_To_ZipCode = "Bill_To_ZipCode";

            Shp_Name_OBL = "Ship Name OBL";
            Shp_Addr_Line_OBL = "Ship Addr OBL";
            Shp_City_OBL = "Ship City OBL";
            Shp_State_OBL = "Ship State OBL";
            Shp_ZipCode_OBL = "Ship Zip OBL";

            Shp_Code_PKP = "ShipperCode_Pickup";
            ShipperDetails_PKP = "ShipperDetails_Pickup";

            billingStatus = string.Empty;
            ShipperList = new Dictionary<string, string>();
            ConsigneeList = new Dictionary<string, clsCnCWord>();

            arValidClasscode = new ArrayList();
            ClassCodesList();
        }


        
        public void StaticVariable_Release()
        {
            

            billingStatus = null;
            cryptographyKey = null;
            processLogfileName = null;
            processLogFilePath = null;

            executionServerName = null;
            executionServiceName = null;
            destinationServerDetailsList = null;


            processLogWriter = null;

            if (dtfieldInfo != null)
            {
                dtfieldInfo.Dispose();
                dtfieldInfo = null;
            }

            //if (dt_HeaderFileddata != null)
            //{
            //    dt_HeaderFileddata.Dispose();
            //    dt_HeaderFileddata = null;
            //}
            //if (dt_lineitemdata != null)
            //{
            //    dt_lineitemdata.Dispose();
            //    dt_lineitemdata = null;
            //}

            if (dt_SortedHeaderFielddata != null)
            {
                dt_SortedHeaderFielddata.Dispose();
                dt_SortedHeaderFielddata = null;
            }
            if (dt_SortedLineitemdata != null)
            {
                dt_SortedLineitemdata.Dispose();
                dt_SortedLineitemdata = null;
            }


            if (iQPUBLIC.PublicComponents.dtDataExchange != null)
            {
                iQPUBLIC.PublicComponents.dtDataExchange.Dispose();
                iQPUBLIC.PublicComponents.dtDataExchange = null;
            }
           


            iQPUBLIC.PublicComponents.htMyVariable = null;


            //DBUtilities.appConnectionString = null;

            //DBUtilities.appConnectionString = null;
            DBUtilities.prodConnectionString = null;
            DBUtilities.DB2ConnectionString = null;

            if (DBUtilities.db2Connection != null)
            {
                DBUtilities.db2Connection.Dispose();
            }

            if (DBUtilities.sqlConnection != null)
            {
                DBUtilities.sqlConnection.Dispose();
            }

            objExtractedData = null;

            ShipperList = null;
            ConsigneeList = null;
            //objCNC.Dispose();
            //objCNC = null;

            arValidClasscode = null;
        }

        private bool GetAppConfigData()
        {
            bool status = false;
            try
            {
                string appConfigPath = ConfigurationManager.AppSettings["AppConfigPath"].ToString();
                if (!string.IsNullOrEmpty(appConfigPath))
                {
                    if (File.Exists(appConfigPath))
                    {
                        string connectionString = File.ReadAllText(appConfigPath);
                        if (!string.IsNullOrEmpty(connectionString) && !string.IsNullOrEmpty(cryptographyKey))
                        {
                            DBUtilities.appConnectionString = Cryptography.CryptographyOperation.Decrypt(connectionString.Trim(), cryptographyKey);
                            status = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
            return status;
        }

        public void ClassCodesList()
        {
            arValidClasscode.Add("50");
            arValidClasscode.Add("55");
            arValidClasscode.Add("60");
            arValidClasscode.Add("65");
            arValidClasscode.Add("70");
            arValidClasscode.Add("77");
            arValidClasscode.Add("77.5");
            arValidClasscode.Add("85");
            arValidClasscode.Add("92");
            arValidClasscode.Add("92.5");
            arValidClasscode.Add("100");
            arValidClasscode.Add("110");
            arValidClasscode.Add("125");
            arValidClasscode.Add("150");
            arValidClasscode.Add("175");
            arValidClasscode.Add("200");
            arValidClasscode.Add("250");
            arValidClasscode.Add("300");
            arValidClasscode.Add("400");
            arValidClasscode.Add("500");
        }
        public void StartDataExtractionService()
        {

            try
            {
                StaticVariable_Initialize();

                DateTime startTime = DateTime.Now;

                ConfigurationManager.RefreshSection("connectionStrings");


                billingStatus = ConfigurationManager.AppSettings["BillingStatus"].ToString();

                DBUtilities.DB2ConnectionString = ConfigurationManager.ConnectionStrings["AS400DBConnString"].ToString();

                if (!GetAppConfigData())
                {
                    return;
                }

                executionServerName = Environment.MachineName;


                if (string.IsNullOrEmpty(executionServerName))
                {
                    InsertErrorLog(0, ErrorType.S, "Server/Machine name not retrieved", executionServiceName, "StartDataExtractionService"); // DocumentID as 0
                    return;
                }


                if (!string.IsNullOrEmpty(DBUtilities.appConnectionString))
                {
                    DataTable appSettingTable = DBUtilities.GetDataTableFromSP(DBUtilities.appConnectionString, "USP_Appsetting_Getdata").Response;
                    if (!ValidateAppSettingData(appSettingTable))
                    {
                        return;
                    }


                    //Validate Execution Server Name present in AppSetting
                    AppSetting setting = new AppSetting();
                    setting = destinationServerDetailsList.Where(a => a.SettingName == executionServerName).FirstOrDefault();
                    if (setting == null)
                    {
                        InsertErrorLog(0, ErrorType.S, "Execution Server/Machine name not present in appsetting data", executionServiceName, "StartDataExtractionService"); // DocumentID as 0
                        //WriteLog("Execution Server/Machine name not present in appsetting data. ", LogType.E);
                        return;
                    }

                    if (appSettingTable != null)
                    {
                        appSettingTable.Dispose();
                        appSettingTable = null;
                    }

                    setting = null;

                    
                }


                if (!string.IsNullOrEmpty(DBUtilities.appConnectionString))
                {
                    dtfieldInfo = DBUtilities.GetDataTableFromSP(DBUtilities.appConnectionString, "USP_FieldInfo_KM2_Getdata").Response;

                    if (dtfieldInfo == null)
                    {

                        InsertErrorLog(0, ErrorType.S, "fieldInfo table not retrieved from database ", executionServiceName, "StartDataExtractionService"); // DocumentID as 0
                        //WriteLog("fieldInfo table not retrieved from database. ", LogType.E);
                        return;
                    }
                }

                DBUtilities.appConnectionString = null;

                if (string.IsNullOrEmpty(DBUtilities.prodConnectionString))
                {
                    InsertErrorLog(0, ErrorType.S, "ProductionServer Connetion was Empty", executionServiceName, "StartDataExtractionService"); // DocumentID as 0
                    //WriteLog("SQL ProductionServer Connetion was Empty", LogType.E);
                    return;
                }

                if (string.IsNullOrEmpty(DBUtilities.DB2ConnectionString))
                {

                    InsertErrorLog(0, ErrorType.S, "DB2 ProductionServer Connetion was Empty", executionServiceName, "StartDataExtractionService"); // DocumentID as 0
                   // WriteLog("DB2 ProductionServer Connetion was Empty", LogType.E);
                    return;
                }


                DataTable documentMasterTable = GetDocumentMasterData();

                if (documentMasterTable != null && documentMasterTable.Rows.Count > 0)
                {
                    //Create Process Log
                    processLogWriter = CommonModule.Common.CreateProcessLogFile(processLogFilePath, processLogfileName);
                    WriteLog("Application Process Log File Created", LogType.P);
                    WriteLog("Document Master Table Data Fetched and Count is: " + documentMasterTable.Rows.Count, LogType.P);

                    bool result = PerformDataExtraction_KM2(documentMasterTable);

                    if (result)
                    {
                        WriteLog("DataExtraction Service Process Completed and Count is: " + documentMasterTable.Rows.Count, LogType.P);
                    }

                    processLogWriter.DisposeProcessLog();
                }


                if (documentMasterTable != null)
                {
                    documentMasterTable.Dispose();
                }
                
                documentMasterTable = null;
                
            }
            catch (Exception ex)
            {
                InsertErrorLog(0, ErrorType.S, ex.Message, executionServiceName, "StartDataExtractionService"); // DocumentID as 0
                //WriteLog("StartDataExtractionService: Error: " + ex.StackTrace.ToString(), LogType.E);
            }
            finally
            {
                UpdateServiceHealthCheck();
                StaticVariable_Release();
                //processLogWriter.DisposeProcessLog();
            }
        }

        public static bool ValidateAppSettingData(DataTable appSettingTable)
        {
            bool result = false;

            string server = string.Empty;
            string database = string.Empty;
            string username = string.Empty;
            string password = string.Empty;

            try
            {
                if (appSettingTable != null && appSettingTable.Rows.Count > 0)
                {
                    destinationServerDetailsList = (from p in appSettingTable.AsEnumerable()
                                                    where p.Field<string>("SuccessOrFail") == "S" && p.Field<string>("SettingHeader") == "TruCap Server(s)"
                                                    select new AppSetting
                                                    {
                                                        SettingName = p.Field<string>("SettingName"),
                                                        SettingValue = p.Field<string>("SettingValue")
                                                    }).ToList();
                    processLogFilePath = appSettingTable.AsEnumerable().Where(a => a.Field<string>("SuccessOrFail") == "S" && a.Field<string>("SettingName") == "LogFilePath" && a.Field<string>("SettingHeader") == "Error Log File Path").Select(a => a.Field<string>("SettingValue")).FirstOrDefault();
                    server = appSettingTable.AsEnumerable().Where(a => a.Field<string>("SuccessOrFail") == "S" && a.Field<string>("SettingName") == "Server" && a.Field<string>("SettingHeader") == "Database details - Production").Select(a => a.Field<string>("SettingValue")).FirstOrDefault();
                    database = appSettingTable.AsEnumerable().Where(a => a.Field<string>("SuccessOrFail") == "S" && a.Field<string>("SettingName") == "Database" && a.Field<string>("SettingHeader") == "Database details - Production").Select(a => a.Field<string>("SettingValue")).FirstOrDefault();
                    username = appSettingTable.AsEnumerable().Where(a => a.Field<string>("SuccessOrFail") == "S" && a.Field<string>("SettingName") == "Username" && a.Field<string>("SettingHeader") == "Database details - Production").Select(a => a.Field<string>("SettingValue")).FirstOrDefault();
                    password = appSettingTable.AsEnumerable().Where(a => a.Field<string>("SuccessOrFail") == "S" && a.Field<string>("SettingName") == "Password" && a.Field<string>("SettingHeader") == "Database details - Production").Select(a => a.Field<string>("SettingValue")).FirstOrDefault();


                    DBUtilities.prodConnectionString = "Data Source=" + server + ";Initial Catalog=" + database + ";User ID=" + username + ";Password=" + password + ";Connect Timeout=30";
                    //DBUtilities.prodConnectionString = "Data Source=" + server + ",1433" + ";Initial Catalog=" + database + ";User ID=" + username + ";Password=" + password + ";Connect Timeout=30";
                    if (destinationServerDetailsList != null && destinationServerDetailsList.Count > 0 && !string.IsNullOrEmpty(processLogFilePath) && !string.IsNullOrEmpty(server) && !string.IsNullOrEmpty(database) && !string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                    {
                        result = true;
                    }
                    if (destinationServerDetailsList != null && destinationServerDetailsList.Count > 0 && !string.IsNullOrEmpty(processLogFilePath))
                    {
                        result = true;
                    }


                }
            }
            catch (Exception ex)
            {
                WriteLog("Appsetting data validation failed. Service exited." + ex.Message, LogType.E);
                InsertErrorLog(0, ErrorType.S, "Appsetting data validation failed:" + ex.Message, executionServiceName, "ValidateAppSettingData"); // DocumentID as 0
                return false;
            }
            finally
            {
                server = null;
                database = null;
                username = null;
                password = null;
            }
            return result;
        }
        public bool PerformDataExtraction_KM2(DataTable documentMasterTable)
        {
            bool result = false;
            long documentID = 0;

            string inputFileName = string.Empty;
            string outputFileName = string.Empty;
            string documentName = string.Empty;
            string preprocessingFilePath = string.Empty;
            //string DataExtraction_HeaderFieldData = string.Empty;

            //DateTime startTime = DateTime.Now;
            try
            {

                DataExchange.CreateStructure();


                objKM2F3 = new SAIABOL01IQ.clsF3();
                objKM2F5 = new SAIABOL01IQ.clsF5();
                objKM2F7 = new SAIABOL01IQ.clsF7();
                objKM2F27 = new SAIABOL01IQ.clsF27();

                objKM2F6 = new SAIABOL01IQ.clsF6();
                objKM2F23 = new SAIABOL01IQ.clsF23();
                objKM2F24 = new SAIABOL01IQ.clsF24();

                //RetStructF3 retStructF3Value = new RetStructF3();
                //retStructF3Value.Status = "F";
                //retStructF3Value.NoOfieldsSuspects = 0;

                int intCurrPageNumber = 1;



                foreach (DataRow document in documentMasterTable.AsEnumerable())
                {
                    documentID = 0;
                    DateTime startTime = DateTime.Now;
                    PageCount = 0;

                   

                    try
                    {
                        documentID = Convert.ToInt64(document["DocumentID"]);

                        WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                        WriteLog("DataExtraction process started for document : " + documentID, LogType.P);


                        //WriteLog("Start connection string intilization  ", LogType.P);
                        //*****************Add ConnectionString in iQPUBLIC.PublicComponents.htMyVariable
                        if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_SQLConn") == false)
                        {
                            iQPUBLIC.PublicComponents.htMyVariable.Add("ConnectionString_SQLConn", DBUtilities.prodConnectionString);
                        }

                        if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_DB2Conn") == false)
                        {
                            iQPUBLIC.PublicComponents.htMyVariable.Add("ConnectionString_DB2Conn", DBUtilities.DB2ConnectionString);
                        }
                        //**************************************************************************


                        DBUtilities.db2Connection.ConnectionString = DBUtilities.DB2ConnectionString;
                        DBUtilities.sqlConnection.ConnectionString = DBUtilities.prodConnectionString;

                        DBUtilities.SQLConn_Open();
                        DBUtilities.DB2Conn_Open();


                        documentName = document["FileName"].ToString();


                        inputFileName = document["DestinationPath"].ToString(); 
                        preprocessingFilePath = Path.GetDirectoryName(inputFileName) + "\\Preprocessed";  

                        iQPUBLIC.PublicComponents.InputBatchPath = preprocessingFilePath;
                        iQPUBLIC.PublicComponents.PrimaryImagePath = preprocessingFilePath + "\\" + documentName;

                        inputFileName = preprocessingFilePath + "\\" + Path.GetFileNameWithoutExtension(documentName) + ".PRO";
                        iQPUBLIC.PublicComponents.PROPath = inputFileName;

                        WriteLog("ImagePROPath: " + inputFileName, LogType.P);

                        if (File.Exists(inputFileName))
                        {
                            intCurrPageNumber = 1;

                            if (!DBNull.Value.Equals(document["BOLPageNo"]))
                            {
                                if (document["BOLPageNo"].ToString().Trim() != "")
                                {
                                    intCurrPageNumber = Convert.ToInt32(document["BOLPageNo"].ToString().Trim().Split(',')[0]);
                                }
                            }


                            Match matchResult = Regex.Match(documentName, "[\\d]+");
                            if (matchResult.Success == true)
                            {
                                if (GetBillingStatus(matchResult.Value) == true) //Continue if FHSTAT = 'PP' Only
                                {
                                    if (GetEDIStatus(matchResult.Value) == false)  //Skip EDI Documnet
                                    {
                                        WriteLog("Start: LoadPRO: " + documentName, LogType.P);
                                        if (LoadPRO(documentID, documentName, inputFileName, intCurrPageNumber) == true)
                                        {
                                            WriteLog("End: LoadPRO: " + documentName, LogType.P);
                                            iQPUBLIC.PublicComponents.dtDataExchange = DataExchange.dtDataExg;
                                            DataExtraction_BOL_KM2(documentID, documentName, intCurrPageNumber);
                                        }
                                    }
                                    else
                                    {
                                        WriteLog(documentName + ":" + " EDI Document", LogType.E);
                                        //InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " EDI Document", executionServiceName, "PerformDataExtraction_KM2");
                                        InsertErrorLog(documentID, ErrorType.S, "EDI Document", executionServiceName, "PerformDataExtraction_KM2");
                                    }
                                }
                                else
                                {
                                    WriteLog(documentName + ":" + " Billed in AS400", LogType.E);
                                    // InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " Billed in AS400", executionServiceName, "PerformDataExtraction_KM2");
                                    InsertErrorLog(documentID, ErrorType.S,  "Billed in AS400", executionServiceName, "PerformDataExtraction_KM2");
                                }
                            }
                            else
                            {
                                WriteLog(documentName + ":" + " Invalid PRO number", LogType.E);
                                InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " Invalid PRO number", executionServiceName, "PerformDataExtraction_KM2");
                            }


                        }
                        else
                        {
                            WriteLog(documentName + ":" + "Not exist for DataExtraction", LogType.E);
                            InsertErrorLog(documentID, ErrorType.S, documentName + ":" + "Not exist for DataExtraction", executionServiceName, "PerformDataExtraction_KM2");
                        }

                        DateTime endTime = DateTime.Now;

                        WriteLog("DataExtraction process completed for document : " + documentID, LogType.P);
                        WriteLog("Total Time Taken for document: " + documentID + ": " + CommonModule.Common.CalculateTotalTime(startTime, endTime) + " PageCount: " + PageCount, LogType.P);

                        //WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);

                    }
                    catch (Exception ex1)
                    {
                        WriteLog(documentName + ":" + " PerformDataExtraction_KM2.: Error: " + ex1.Message, LogType.E);
                        InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " PerformDataExtraction_KM2.: Error: " + ex1.Message, executionServiceName, "PerformDataExtraction_KM2");
                       // UpdateDocumentMaster_ErrorFlag(documentID, startTime, DateTime.Now);

                    }
                    finally
                    {

                        iQPUBLIC.PublicComponents.dtDataExchange = null;
                        DataExchange.ClearData();

                        DBUtilities.DisposeConnection_DB2(DBUtilities.db2Connection);


                        DBUtilities.DisposeConnection_SQL(DBUtilities.sqlConnection);

                        WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);

                    }
                }



                result = true;

            }
            catch (Exception ex)
            {
                WriteLog(documentName + ":" + " PerformDataExtraction_KM2: Error: " + ex.Message, LogType.E);
                InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " PerformDataExtraction_KM2: Error: " + ex.Message, executionServiceName, "PerformDataExtraction_KM2");
                //UpdateDocumentMaster_ErrorFlag(documentID, DateTime.Now, DateTime.Now);
                result = false;
            }

            finally
            {
                objKM2F3 = null;
                objKM2F5 = null;
                objKM2F7 = null;
                objKM2F27 = null;

                objKM2F6 = null;
                objKM2F23 = null;
                objKM2F24 = null;


                inputFileName = null;
                outputFileName = null;
                documentName = null;
                preprocessingFilePath = null;

            }

            return result;
        }

        private bool LoadPRO(long documentID, string documentName, string filePath, int intCurrPageNumber)
        {
            bool result = false;

            try
            {
               // imgOrentation = 0;

                var fi = new FileInfo(filePath);

                if (fi.Length == 0)
                {
                    WriteLog(documentName + ": " + " LoadPRO: MetaData Application Generated 0 KB PRO File.", LogType.E);
                    InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " LoadPRO: MetaData Application Generated 0 KB PRO File.", executionServiceName, "LoadPRO");
                    result = false;
                }
                else
                {
                    objCNC = new ClsCNC();

                    if (objCNC.LoadPROFile(filePath, false, "ICR") > 0)
                    {
                        
                        iQPUBLIC.PublicComponents.oMetaData1 = objCNC.GetMetaData(1);
                        iQPUBLIC.PublicComponents.oMetaData2 = objCNC.GetMetaData(2);

                        if (iQPUBLIC.PublicComponents.oMetaData1 == null || iQPUBLIC.PublicComponents.oMetaData2 == null)
                        {
                            WriteLog(documentName + ": " + "LoadPRO: Error: MetaData Return Null Value ", LogType.E);
                            InsertErrorLog(documentID, ErrorType.S, documentName + ": " + "LoadPRO: Error: MetaData Return Null Value", executionServiceName, "LoadPRO");
                            result = false;
                        }

                        PageCount = objCNC.GetMetaData(1).PageCount;

                        if (PageCount < 10)
                        {

                            iQPUBLIC.PublicComponents.ImageHeight = iQPUBLIC.PublicComponents.oMetaData1.Page[intCurrPageNumber].ImageHeight;
                            iQPUBLIC.PublicComponents.ImageWidth = iQPUBLIC.PublicComponents.oMetaData1.Page[intCurrPageNumber].ImageWidth;



                            if (iQPUBLIC.PublicComponents.htMyVariable.Contains("rootObject") == true)
                            {
                                iQPUBLIC.PublicComponents.htMyVariable.Remove("rootObject");
                            }
                            iQPUBLIC.PublicComponents.htMyVariable.Add("rootObject", objCNC);

                            result = true;
                        }
                        else
                        {
                            //***********PageCount >10 No Extraction 
                            WriteLog(documentName + ": " + " LoadPRO: PageCount: " + Convert.ToString(PageCount), LogType.E);
                            InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " LoadPRO: PageCount: " + Convert.ToString(PageCount), executionServiceName, "LoadPRO");
                            result = false;
                        }

                    }
                    else
                    {
                        WriteLog(documentName + ": " + " LoadPRO: Error: Application unable to read document information ", LogType.E);
                        InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " LoadPRO: Error: Application unable to read document information ", executionServiceName, "LoadPRO");
                        result = false;
                    }
                }
            }
            catch (Exception ex)
            {

                if (iQPUBLIC.PublicComponents.oMetaData1 != null && iQPUBLIC.PublicComponents.oMetaData1.PageCount < intCurrPageNumber)
                {
                    WriteLog(documentName + ":" + " LoadPRO: MetaData Application Generated 0 KB PRO for page "+ intCurrPageNumber, LogType.E);
                    InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " LoadPRO: MetaData Application Generated 0 KB PRO for page " + intCurrPageNumber, executionServiceName, "LoadPRO");
                }
                else
                {
                    WriteLog(documentName + ":" + " LoadPRO: Error: " + ex.Message, LogType.E);
                    InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " LoadPRO: Error: " + ex.Message, executionServiceName, "LoadPRO");
                }
                
                result = false;
            }
            finally
            {
                if (objCNC != null)
                {
                    objCNC.Dispose();
                    objCNC = null;
                }
            }
            return result;
        }

        private void DataExtraction_BOL_KM2(long documentID, string documentName, int intCurrPageNumber)
        {
            string DataExtraction_HeaderFieldData = string.Empty;
            string DataExtraction_LineFieldData = string.Empty;

            StringBuilder objDataExtraction_HeaderField = new StringBuilder();
            StringBuilder objDataExtraction_LineField = new StringBuilder();

            clsCncMetaData ObjMetaData = null;
            DateTime startTime = DateTime.Now;
            DateTime startTime_Line = DateTime.Now;

            //int _intCurrentFieldId = 0;
            string _CurrentFieldName = "";

            try
            {

                if (dt_SortedHeaderFielddata != null && dt_SortedHeaderFielddata.Rows.Count > 0) { dt_SortedHeaderFielddata.Clear(); }
                if (dt_SortedLineitemdata != null && dt_SortedLineitemdata.Rows.Count > 0) { dt_SortedLineitemdata.Clear(); }



                objDataExtraction_HeaderField.Append("<?xml version=" + "'1.0'" + "?>");
                objDataExtraction_HeaderField.Append("<HeaderDetails>");

                objDataExtraction_LineField.Append("<?xml version=" + "'1.0'" + "?>");
                objDataExtraction_LineField.Append("<RowDetails>");

                string strFieldVal = "H";

                bool isF3Invoked = false;
                string tempSHP_Code = "*****";
                string _BOLSET2_ShipperValidFlag = string.Empty;
                string ConsigneeDetails_Pickup = string.Empty;
                bool consigMiscFlag = true;

                int imgOrentation = 0;
                imgOrentation = iQPUBLIC.PublicComponents.oMetaData1.Page[intCurrPageNumber].Orentation;

                //clsCnCWord objword_ConsCode = new clsCnCWord();



                StringBuilder objDataExtraction_ConsCode = new StringBuilder();

                Cons_Code_Misc = "";

                SrNo = 0;
                Shp_ChangeGreenToRed = false;
                Con_ChangeGreenToRed = false;
                BillTo_ChangeGreenToRed = false;

                ShipperList.Clear();
                ConsigneeList.Clear();

                foreach (DataRow drfieldInfo in dtfieldInfo.AsEnumerable())
                {
                    _CurrentFieldName = "";
                    isF3Invoked = false;

                    iQPUBLIC.PublicComponents.intCurrentFieldId = Convert.ToInt32(drfieldInfo["FieldID"].ToString());
                    iQPUBLIC.PublicComponents.CurrentFieldName = drfieldInfo["SysFieldName"].ToString();

                    _CurrentFieldName = iQPUBLIC.PublicComponents.CurrentFieldName;


                    if (drfieldInfo["SysFieldName"].ToString().Trim().ToLower().StartsWith("line1"))
                    {
                        startTime_Line = DateTime.Now;
                        
                        if (PublicComponents.htMyVariable.ContainsKey("BOLSET4_ShpCode") == true)
                        {
                            PublicComponents.htMyVariable.Remove("BOLSET4_ShpCode");
                        }
                        PublicComponents.htMyVariable.Add("BOLSET4_ShpCode", tempSHP_Code);


                        if (PublicComponents.htMyVariable.ContainsKey("BOLSET2_ShipperValidFlag") == true)
                        {
                            PublicComponents.htMyVariable.Remove("BOLSET2_ShipperValidFlag");
                        }
                        PublicComponents.htMyVariable.Add("BOLSET2_ShipperValidFlag", _BOLSET2_ShipperValidFlag);
                    }

                    if (drfieldInfo["Active"].ToString() != "True")
                    {
                        continue;
                    }

                    if (drfieldInfo["FieldStructure"].ToString() == "1")
                    {
                        ObjMetaData = iQPUBLIC.PublicComponents.oMetaData1;
                    }
                    else
                    {
                        ObjMetaData = iQPUBLIC.PublicComponents.oMetaData2;
                    }

                    retStructF3Value.Status = "F";
                    retStructF3Value.NoOfieldsSuspects = 0;
                    retStructF3Value.Words = null;

                    retStructF27Value.Status = "F";
                    retStructF27Value.ContentsOfField = "";



                    if (drfieldInfo["F24RoutineNo"].ToString().Trim() != "")
                    {
                        DataExchange.UpdateRecBeforeCall(documentName, drfieldInfo["FieldID"].ToString(), drfieldInfo["DispFieldName"].ToString(), strFieldVal, "F24", drfieldInfo["F24RoutineNo"].ToString(), drfieldInfo["F24Arguments"].ToString(), intCurrPageNumber);
                        RetStructF24 retStruct24Value = objKM2F24.F24(drfieldInfo["F24RoutineNo"].ToString(), drfieldInfo["F24Arguments"].ToString());
                        DataExchange.UpdateRecAfterCall(documentName, drfieldInfo["FieldID"].ToString(), strFieldVal, "F24", drfieldInfo["F24RoutineNo"].ToString(), retStruct24Value.Status, retStruct24Value);
                    }

                    if (drfieldInfo["F27RoutineNo"].ToString().Trim() != "")
                    {
                        isF3Invoked = false;
                        DataExchange.UpdateRecBeforeCall(documentName, drfieldInfo["FieldID"].ToString(), drfieldInfo["DispFieldName"].ToString(), strFieldVal, "F27", drfieldInfo["F27RoutineNo"].ToString(), drfieldInfo["F27Arguments"].ToString(), intCurrPageNumber);
                        retStructF27Value = objKM2F27.F27(drfieldInfo["F27RoutineNo"].ToString(), drfieldInfo["F27Arguments"].ToString(), drfieldInfo["FieldID"].ToString(), "", "H");
                        DataExchange.UpdateRecAfterCall(documentName, drfieldInfo["FieldID"].ToString(), strFieldVal, "F27", drfieldInfo["F27RoutineNo"].ToString(), retStructF27Value.Status, retStructF27Value);

                    }
                    else if (drfieldInfo["F3RoutineNo"].ToString().Trim() != "")
                    {

                        isF3Invoked = true;
                        DataExchange.UpdateRecBeforeCall(documentName, drfieldInfo["FieldID"].ToString(), drfieldInfo["DispFieldName"].ToString(), strFieldVal, "F3", drfieldInfo["F3RoutineNo"].ToString(), drfieldInfo["F3Arguments"].ToString(), intCurrPageNumber);
                        retStructF3Value = objKM2F3.F3(ObjMetaData, intCurrPageNumber, drfieldInfo["F3RoutineNo"].ToString(), drfieldInfo["F3Arguments"].ToString());
                        DataExchange.UpdateRecAfterCall(documentName, drfieldInfo["FieldID"].ToString(), strFieldVal, "F3", drfieldInfo["F3RoutineNo"].ToString(), retStructF3Value.Status, retStructF3Value);

                    }


                    if (drfieldInfo["F5RoutineNo"].ToString().Trim() != "")
                    {
                        if (retStructF3Value.Status == "S")
                        {
                            DataExchange.UpdateRecBeforeCall(documentName, drfieldInfo["FieldID"].ToString(), drfieldInfo["DispFieldName"].ToString(), strFieldVal, "F5", drfieldInfo["F5RoutineNo"].ToString(), drfieldInfo["F5Arguments"].ToString(), intCurrPageNumber);
                            RetStructF5 retStructF5Value = objKM2F5.F5(retStructF3Value.Words[0], drfieldInfo["F5RoutineNo"].ToString(), drfieldInfo["F5Arguments"].ToString());

                            if (retStructF5Value.Status == "S")
                            {
                                if (retStructF5Value.Words != null)
                                {
                                    retStructF3Value.Words[0].strWord = retStructF5Value.Words.strWord;
                                }
                            }
                           
                            DataExchange.UpdateRecAfterCall(documentName, drfieldInfo["FieldID"].ToString(), strFieldVal, "F5", drfieldInfo["F5RoutineNo"].ToString(), retStructF5Value.Status, retStructF5Value);
                        }
                    }

                    if (drfieldInfo["F7RoutineNo"].ToString().Trim() != "")
                    {
                        if (retStructF3Value.Status == "S" || retStructF27Value.Status == "S")
                        {
                            DataExchange.UpdateRecBeforeCall(documentName, drfieldInfo["FieldID"].ToString(), drfieldInfo["DispFieldName"].ToString(), strFieldVal, "F7", drfieldInfo["F7RoutineNo"].ToString(), drfieldInfo["F7Arguments"].ToString(), intCurrPageNumber);
                            RetStructF7 retStructF7Value;
                            if (retStructF27Value.Status == "S")
                            {
                                if (retStructF27Value.ContentsOfField != null)
                                {
                                    retStructF7Value = objKM2F7.F7(retStructF27Value.ContentsOfField, drfieldInfo["F7RoutineNo"].ToString(), drfieldInfo["F7Arguments"].ToString());
                                }
                            }
                            else if (retStructF3Value.Status == "S")
                            {
                                if (retStructF3Value.Words != null)
                                {
                                    retStructF7Value = objKM2F7.F7(retStructF3Value.Words[0].strWord, drfieldInfo["F7RoutineNo"].ToString(), drfieldInfo["F7Arguments"].ToString());
                                }
                            }

                            if (retStructF7Value.Status == "F")
                            {
                                retStructF27Value.Status = "F";
                                retStructF27Value.ContentsOfField = "";
                                retStructF3Value.Status = "F";
                                retStructF3Value.NoOfieldsSuspects = 0;
                                retStructF3Value.Words = null;
                            }

                            DataExchange.UpdateRecAfterCall(documentName, drfieldInfo["FieldID"].ToString(), strFieldVal, "F7", drfieldInfo["F7RoutineNo"].ToString(), retStructF7Value.Status, retStructF7Value);
                        }
                    }

                    if (drfieldInfo["F6RoutineNo"].ToString().Trim() != "")
                    {
                        if (retStructF3Value.Status == "S")
                        {
                            DataExchange.UpdateRecBeforeCall(documentName, drfieldInfo["FieldID"].ToString(), drfieldInfo["DispFieldName"].ToString(), strFieldVal, "F6", drfieldInfo["F6RoutineNo"].ToString(), drfieldInfo["F6Arguments"].ToString(), intCurrPageNumber);
                            RetStructF6 retStructF6Value = objKM2F6.F6(retStructF3Value.Words[0], drfieldInfo["F6RoutineNo"].ToString(), drfieldInfo["F6Arguments"].ToString());

                            DataExchange.UpdateRecAfterCall(documentName, drfieldInfo["FieldID"].ToString(), strFieldVal, "F6", drfieldInfo["F6RoutineNo"].ToString(), retStructF6Value.Status, retStructF6Value);
                        }
                    }

                    if (drfieldInfo["F23RoutineNo"].ToString().Trim() != "")
                    {
                        DataExchange.UpdateRecBeforeCall(documentName, drfieldInfo["FieldID"].ToString(), drfieldInfo["DispFieldName"].ToString(), "H", "F23", drfieldInfo["F23RoutineNo"].ToString(), drfieldInfo["F23Arguments"].ToString(), intCurrPageNumber);
                        RetStructF23 retStruct23Value = objKM2F23.F23(strFieldVal, drfieldInfo["F23RoutineNo"].ToString(), drfieldInfo["F23Arguments"].ToString());
                        DataExchange.UpdateRecAfterCall(documentName, drfieldInfo["FieldID"].ToString(), "H", "F23", drfieldInfo["F23RoutineNo"].ToString(), retStruct23Value.Status, retStruct23Value);
                    }

                    if (drfieldInfo["SysFieldName"].ToString().Trim().ToLower() == "line1")
                    {
                        DataTable dtLineItems = null;
                        if (PublicComponents.htMyVariable.ContainsKey("BOLSET4_LineItemValue") == true)
                        {
                            dtLineItems = (DataTable)iQPUBLIC.PublicComponents.htMyVariable["BOLSET4_LineItemValue"];
                        }

                        SrNo = 0;

                        if (dtLineItems == null)
                        {

                            objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Quantity", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                            objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Class_Code", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                            objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "HazMat_Code", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                            objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Pallate_Code", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                            objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Description", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                            objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "NMFC", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                            objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Weight", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                        }
                        else
                        {
                            if (dtLineItems.Rows.Count > 0)
                            {
                                dtLineItems.DefaultView.Sort = "PROLineNo";

                                dtLineItems = dtLineItems.DefaultView.ToTable();

                                for (int irowNo = 0; irowNo < dtLineItems.Rows.Count; irowNo++)
                                {
                                    // clsCnCWord objLineWord;
                                    if (!DBNull.Value.Equals(dtLineItems.Rows[irowNo]["WQuantity"]))
                                    {
                                        // clsCnCWord objLineWord_Quantity = (clsCnCWord)dtLineItems.Rows[irowNo]["WQuantity"];

                                        // objLineWord = (clsCnCWord)dtLineItems.Rows[irowNo]["WQuantity"];
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField((clsCnCWord)dtLineItems.Rows[irowNo]["WQuantity"], "Quantity", documentID, irowNo + 1, true, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                        //objLineWord_Quantity = null;
                                    }
                                    else
                                    {
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Quantity", documentID, irowNo + 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));

                                    }

                                    if (!DBNull.Value.Equals(dtLineItems.Rows[irowNo]["WClass Code"]))
                                    {
                                        // objLineWord = (clsCnCWord)dtLineItems.Rows[irowNo]["WClass Code"];
                                        //********************Added on 10Nov21 Replace .0 and .00 and accept only valid class codes************
                                        clsCnCWord clsClassCode = (clsCnCWord)dtLineItems.Rows[irowNo]["WClass Code"];
                                        string strWord = Class_CodeValidate(clsClassCode.strWord);
                                        if (arValidClasscode.Contains(strWord) == true)
                                        {
                                            clsClassCode.strWord = strWord;
                                            clsClassCode.Flag = "1";
                                        }
                                        else
                                        {
                                            clsClassCode.Flag = "0";
                                        }
                                        //*******************************************************
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField((clsCnCWord)dtLineItems.Rows[irowNo]["WClass Code"], "Class_Code", documentID, irowNo + 1, true, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }
                                    else
                                    {
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Class_Code", documentID, irowNo + 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }

                                    if (!DBNull.Value.Equals(dtLineItems.Rows[irowNo]["WHazMat Code"]))
                                    {
                                        // objLineWord = (clsCnCWord)dtLineItems.Rows[irowNo]["WHazMat Code"];
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField((clsCnCWord)dtLineItems.Rows[irowNo]["WHazMat Code"], "HazMat_Code", documentID, irowNo + 1, true, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }
                                    else
                                    {
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "HazMat_Code", documentID, irowNo + 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }


                                    if (!DBNull.Value.Equals(dtLineItems.Rows[irowNo]["WPallate Code"]))
                                    {
                                        // clsCnCWord objLineWord_PallateCode = (clsCnCWord)dtLineItems.Rows[irowNo]["WPallate Code"];
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField((clsCnCWord)dtLineItems.Rows[irowNo]["WPallate Code"], "Pallate_Code", documentID, irowNo + 1, true, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }
                                    else
                                    {
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Pallate_Code", documentID, irowNo + 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }



                                    if (!DBNull.Value.Equals(dtLineItems.Rows[irowNo]["WDescription"]))
                                    {
                                        //clsCnCWord objLineWord_Description = (clsCnCWord)dtLineItems.Rows[irowNo]["WDescription"];
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField((clsCnCWord)dtLineItems.Rows[irowNo]["WDescription"], "Description", documentID, irowNo + 1, true, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }
                                    else
                                    {
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Description", documentID, irowNo + 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }

                                    if (!DBNull.Value.Equals(dtLineItems.Rows[irowNo]["WNMFC"]))
                                    {
                                        // clsCnCWord objLineWord_NMFC = (clsCnCWord)dtLineItems.Rows[irowNo]["WNMFC"];
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField((clsCnCWord)dtLineItems.Rows[irowNo]["WNMFC"], "NMFC", documentID, irowNo + 1, true, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }
                                    else
                                    {
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "NMFC", documentID, irowNo + 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }

                                    if (!DBNull.Value.Equals(dtLineItems.Rows[irowNo]["WWeight"]))
                                    {
                                        // clsCnCWord objLineWord_Weight = (clsCnCWord)dtLineItems.Rows[irowNo]["WWeight"];
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField((clsCnCWord)dtLineItems.Rows[irowNo]["WWeight"], "Weight", documentID, irowNo + 1, true, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }
                                    else
                                    {
                                        objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Weight", documentID, irowNo + 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                    }

                                }
                            }
                            else
                            {
                                objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Quantity", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Class_Code", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "HazMat_Code", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Pallate_Code", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Description", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "NMFC", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                                objDataExtraction_LineField.Append(AppendExtractedData_LineField(null, "Weight", documentID, 1, false, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                            }
                        }

                        if (dtLineItems != null)
                        {
                            dtLineItems.Dispose();
                            dtLineItems = null;
                        }

                    }
                    else
                    {

                        if (drfieldInfo["SysFieldName"].ToString().Trim() == "Con_Code")
                        {
                            objDataExtraction_ConsCode = AppendExtractedData_HeaderField_F27(retStructF27Value, drfieldInfo["SysFieldName"].ToString().Trim(), documentID, drfieldInfo["LocalFieldTitle"].ToString().Trim(), consigMiscFlag, ConsigneeDetails_Pickup);
                        }
                        else
                        {

                            if (isF3Invoked == true)
                            {
                                objDataExtraction_HeaderField.Append(AppendExtractedData_HeaderField_F3(retStructF3Value, drfieldInfo["SysFieldName"].ToString().Trim(), documentID, drfieldInfo["LocalFieldTitle"].ToString().Trim()));
                            }
                            else
                            {

                                objDataExtraction_HeaderField.Append(AppendExtractedData_HeaderField_F27(retStructF27Value, drfieldInfo["SysFieldName"].ToString().Trim(), documentID, drfieldInfo["LocalFieldTitle"].ToString().Trim(), consigMiscFlag, ConsigneeDetails_Pickup));

                            }
                        }
                    }


                    if (drfieldInfo["SysFieldName"].ToString().Trim().StartsWith("Shp_Code"))
                    {
                        if (retStructF27Value.ContentsOfField != null)
                        {
                            tempSHP_Code = retStructF27Value.ContentsOfField;
                        }

                    }

                    //****************Added ConMisc Logic on 12Aug21*******
                    if (drfieldInfo["SysFieldName"].ToString().Trim() == "ConsigneeDetails_Pickup")
                    {
                        if (retStructF27Value.ContentsOfField != null)
                        {
                            ConsigneeDetails_Pickup = retStructF27Value.ContentsOfField;
                        }
                    }

                    if (drfieldInfo["SysFieldName"].ToString().Trim() == "ConsigneeDetailsOCR_Matched_MasterDB_Key" || drfieldInfo["SysFieldName"].ToString().Trim() == "ConsigneeDetails OCR_matched_Searchengine")
                    {
                        if (retStructF27Value.ContentsOfField != null)
                        {
                            if (retStructF27Value.ContentsOfField.Trim().ToUpper() == "YES")
                            {
                                consigMiscFlag = false;
                            }
                        }
                    }

                    //if (drfieldInfo["SysFieldName"].ToString().Trim() == "ConsigneeDetailsOCR_Matched_MasterDB_Key")
                    //{
                    //    if (retStructF27Value.ContentsOfField != null)
                    //    {
                    //        if (retStructF27Value.ContentsOfField.Trim().ToUpper() == "YES")
                    //        {
                    //            consigMiscFlag = false;
                    //        }
                    //        else
                    //        {
                    //            consigMiscFlag = true;
                    //        }
                    //    }
                    //    else
                    //    {
                    //        consigMiscFlag = true;
                    //    }
                    //}
                    //if (drfieldInfo["SysFieldName"].ToString().Trim() == "ConsigneeDetails OCR_matched_Searchengine")
                    //{
                    //    if (retStructF27Value.ContentsOfField != null)
                    //    {
                    //        if (retStructF27Value.ContentsOfField.Trim().ToUpper() == "YES")
                    //        {
                    //            consigMiscFlag = false;
                    //        }
                    //        else
                    //        {
                    //            consigMiscFlag = true;
                    //        }
                    //    }
                    //    else
                    //    {
                    //        consigMiscFlag = true;
                    //    }
                    //}

                    //*********************************************

                    if (drfieldInfo["SysFieldName"].ToString().Trim() == "Con_TelNo")
                    {
                        if (PublicComponents.htMyVariable.ContainsKey("BOLSET2_ConsigneeMiscellaneousFlag") == true)
                        {
                            if (PublicComponents.htMyVariable["BOLSET2_ConsigneeMiscellaneousFlag"] != null)
                            {
                                if (Convert.ToString(PublicComponents.htMyVariable["BOLSET2_ConsigneeMiscellaneousFlag"]) == "T")
                                {
                                    objDataExtraction_ConsCode = objDataExtraction_ConsCode.Replace(Cons_Code_Misc, "<FieldValue>000000M</FieldValue>");
                                    objDataExtraction_ConsCode = objDataExtraction_ConsCode.Replace("<FieldColor>0</FieldColor>", "<FieldColor>1</FieldColor>");

                                    //************Update on Datatable*********
                                    dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Con_Code)).ToList().ForEach(D => D.SetField("FieldValue", "000000M"));
                                    dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Con_Code)).ToList().ForEach(D => D.SetField("FieldColor", "1"));
                                    //**************************************
                                   
                                }
                            }
                        }
                        objDataExtraction_HeaderField.Append(objDataExtraction_ConsCode);
                    }


                    if (drfieldInfo["SysFieldName"].ToString().Trim() == "Shp_TelPhone_Number")
                    {
                        if (PublicComponents.htMyVariable.ContainsKey("BOLSET2_ShipperValidFlag") == true)
                        {
                            _BOLSET2_ShipperValidFlag = Convert.ToString(PublicComponents.htMyVariable["BOLSET2_ShipperValidFlag"]);
                        }

                    }


                    objExtractedData = null;
                }

                DateTime endTime = DateTime.Now;

                objDataExtraction_HeaderField.Append("</HeaderDetails>");
                objDataExtraction_LineField.Append("</RowDetails>");

                DataExtraction_HeaderFieldData = Convert.ToString(objDataExtraction_HeaderField);
                DataExtraction_LineFieldData = Convert.ToString(objDataExtraction_LineField);
                if (HeaderFieldData_BulkInsert(documentID, DataExtraction_HeaderFieldData))
                //if (HeaderFieldData_BulkInsert(documentID, DataExtraction_HeaderFieldData) && LineFieldData_BulkInsert(documentID, DataExtraction_LineFieldData))
                {
                    string TotaltimeInSec = CommonModule.Common.CalculateTotalTime_Sec(startTime, startTime_Line) + "~" + CommonModule.Common.CalculateTotalTime_Sec(startTime_Line, endTime) + "~" + PageCount;
                    if (UpdateDocumentMaster(documentID, startTime, endTime, TotaltimeInSec, imgOrentation, false))
                    {
                        WriteLog(documentName + ":" + " Document dataExtracted and status updated as DataExtracted on DB", LogType.P);

                        #region "AS400ColorCoding"

                        if (Shp_ChangeGreenToRed == true)
                        {

                            //************Update on Datatable*********
                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Shp_Code)).ToList().ForEach(D => D.SetField("FieldColor", "0"));
                            //**************************************

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Shp_Name)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Shp_Addr_Line)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Shp_City)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Shp_State)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Shp_ZipCode)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                        }

                        string ConMiscCode = "";
                        if (dt_SortedHeaderFielddata.Select("FieldName = 'Con_Code'").Length > 0)
                        {
                            ConMiscCode = dt_SortedHeaderFielddata.Select("FieldName = 'Con_Code'")[0]["FieldValue"].ToString().Trim();
                            if (dt_SortedHeaderFielddata.Select("FieldName = 'Con_Code'")[0]["FieldColor"].ToString().Contains("1") == false)
                            { Con_ChangeGreenToRed = true; }
                        }

                        if (Con_ChangeGreenToRed == true || ConMiscCode == "000000N")
                        {

                            if (ConMiscCode == "000000N")
                            {
                                dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Con_Code)).ToList().ForEach(D => D.SetField("FieldValue", "000000M"));
                            }
                            else
                            {
                                //************Update on Datatable*********
                                dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Con_Code)).ToList().ForEach(D => D.SetField("FieldColor", "0"));
                                //**************************************
                            }
                            if (ConMiscCode != "000000N")
                            {
                                dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Con_Name)).ToList().ForEach(D => D.SetField("FieldColor", "0"));
                            }

                            if (ConMiscCode != "000000N")
                            {
                                dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Con_Addr_Line)).ToList().ForEach(D => D.SetField("FieldColor", "0"));
                            }

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Con_City)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Con_State)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Con_ZipCode)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Con_TelNo)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                        }


                        //*******************************Added on 31OCT21 if Con_Code is Misc i.e. 000000M then Code will bo blank into AS400 ******************
                        if (ConMiscCode == "000000M" || ConMiscCode == "000000N")
                        {
                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Con_Code)).ToList().ForEach(D => D.SetField("FieldColor", "0"));
                        }
                        //*****************************************************************

                        if (Shp_ChangeGreenToRed == true || Con_ChangeGreenToRed == true)
                        {
                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals("Terms")).ToList().ForEach(D => D.SetField("FieldColor", "0"));
                        }

                        if (BillTo_ChangeGreenToRed == true)
                        {
                            //************Update on Datatable*********
                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Bill_To_Code)).ToList().ForEach(D => D.SetField("FieldColor", "0"));
                            //**************************************

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Bill_To_Name)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Bill_To_Addr_Line)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Bill_To_City)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Bill_To_State)).ToList().ForEach(D => D.SetField("FieldColor", "0"));

                            dt_SortedHeaderFielddata.AsEnumerable().Where(s => Convert.ToString(s["FieldName"]).Equals(Bill_To_ZipCode)).ToList().ForEach(D => D.SetField("FieldColor", "0"));


                        }

                        DataRow[] dr_Pro_No = dt_SortedHeaderFielddata.Select("FieldName = '" + Pro_No + "'");
                        if (dr_Pro_No.Length > 0)
                        {
                            DateTime startTime_ColorCoding = DateTime.Now;
                            string FrieghtBillNumber = dr_Pro_No[0]["FieldValue"].ToString();

                            cls_ReturnStatus _ReturnStatus = cls_SaiaBol.Func_GetDataFor_Freight(dt_SortedHeaderFielddata, dt_SortedLineitemdata, FrieghtBillNumber.Trim(), "", DBUtilities.DB2ConnectionString, billingStatus);

                            DateTime endTime_ColorCoding = DateTime.Now;
                            if (_ReturnStatus.Status != null && _ReturnStatus.Status == "S")
                            {
                                UpdateDocumentStatus_As400Color(documentID, (int)DocumentStatus.Completed, _ReturnStatus.Message, startTime_ColorCoding, endTime_ColorCoding, executionServerName);
                                //"Transfer success..!!";
                                WriteLog(documentName + ":" + " Document insert color coding Success on DB", LogType.P);
                            }
                            else
                            {
                                UpdateDocumentStatus_As400Color(documentID, (int)DocumentStatus.Failed, _ReturnStatus.Message, startTime_ColorCoding, endTime_ColorCoding, executionServerName);

                                InsertErrorLog_As400(documentID, ErrorType.S, _ReturnStatus.Value, executionServiceName, "InsertDataIntoDB2"); // DocumentID as 0

                                //"Transfer failed..!!";
                                WriteLog(documentName + ":" + " Document insert color coding failed on DB", LogType.E);
                            }
                        }
                        else
                        {
                            WriteLog(documentID + ": DataExtraction_BOL_KM2: Error: FrieghtBillNumber Blank ", LogType.E);
                            InsertErrorLog(documentID, ErrorType.S, documentID + ":" + "DataExtraction_BOL_KM2: Error: FrieghtBillNumber Blank ", executionServiceName, "DataExtraction_BOL_KM2");
                        }

                        WriteLog("----------------------------------------------------------------------------------------------------------------", LogType.P);
                        WriteLog("---------------------------------End of DataExtraction & ColorCoding---------------------------------", LogType.P);

                        #endregion

                        #region"Insert Shipper Deatils KM2"
                        if (Shp_ChangeGreenToRed != true)
                        {
                            InsertShipperDetails_KM2(documentID, documentName);
                        }
                        #endregion
                    }
                }

            }
            catch (Exception ex)
            {
                // MessageBox.Show(documentID + ": DataExtraction_BOL_KM2: Error: " + ex.Message + "" _intCurrentFieldId:"+ _intCurrentFieldId + " _CurrentFieldName:"+ _CurrentFieldName);
                WriteLog(documentID + ": DataExtraction_BOL_KM2: Error: " + ex.Message + " For FieldName:" + _CurrentFieldName, LogType.E);
                InsertErrorLog(documentID, ErrorType.S, documentID + ":" + "DataExtraction_BOL_KM2: Error: " + ex.Message + " For FieldName:" + _CurrentFieldName, executionServiceName, "DataExtraction_BOL_KM2");
                //UpdateDocumentMaster_ErrorFlag(documentID, startTime, DateTime.Now);
            }
            finally
            {
                DataExtraction_HeaderFieldData = null;
                DataExtraction_LineFieldData = null;

                objDataExtraction_HeaderField = null;
                objDataExtraction_LineField = null;

                ObjMetaData = null;

            }
        }


        public string Class_CodeValidate(string _ClassCode)
        {
            try
            {

                string[] _Array = null;

                if (!string.IsNullOrEmpty(_ClassCode))
                {
                    string[] ClassCodeArray = { ".0", ".00" };

                    foreach (var item in ClassCodeArray)
                    {
                        if (_ClassCode.Trim().Contains(item))
                        {
                            _Array = _ClassCode.Split('.');
                            _ClassCode = Convert.ToString(_Array[0]);
                            //return _ClassCode;
                            break;
                        }
                    }
                }

                if (_ClassCode.Trim().Contains('.'))
                {
                    _ClassCode = _ClassCode.TrimEnd('0');
                }

            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Error in Class_CodeValidate function--->" + ex, LogType.E);
            }
            return _ClassCode;
        }

        public DataSet GetExtractionDetails(string PRONumber, long documentID)
        {
            DataSet outputModel = new DataSet();
            try
            {
                if (!string.IsNullOrEmpty(DBUtilities.prodConnectionString))
                {
                    string spName = "USP_Get_ExtractionDetails";

                    SqlParameter proNumberParam = new SqlParameter("@PRONumber", PRONumber);

                    SqlParameter[] parameters = new SqlParameter[] { proNumberParam };
                    outputModel = DBUtilities.GetDataTableFromSP_As400(DBUtilities.prodConnectionString, spName, parameters);
                   // return outputModel;
                }
            }
            catch (Exception ex)
            {
                WriteLog(documentID + ": " + " GetExtractionDetails failed:" + ex.Message, LogType.E);
                InsertErrorLog(documentID, ErrorType.S, documentID + ":" + " GetExtractionDetails failed:" + ex.Message, executionServiceName, "GetExtractionDetails");
                return null;
            }

            return outputModel;
        }

        private bool UpdateServiceHealthCheck()
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            bool result = false;
            try
            {
                if (!string.IsNullOrEmpty(DBUtilities.prodConnectionString))
                {
                    string spName = "USP_ServiceHealthCheck_Update";
                    SqlParameter serviceNameParameter = new SqlParameter("@ServiceName", executionServiceName);
                    SqlParameter serverNameParameter = new SqlParameter("@ServerName", executionServerName); // server where service is running
                    SqlParameter lastRunTimeParameter = new SqlParameter("@LastRunTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                    SqlParameter[] parameters = new SqlParameter[] { serviceNameParameter, serverNameParameter, lastRunTimeParameter };
                    serviceResponse = DBUtilities.InsertOrUpdateBySP(DBUtilities.prodConnectionString, spName, parameters);
                    if (serviceResponse.Status == ServiceResponseStatus.Success)
                    {
                        result = true;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("Update To Service Health Check Failed:" + ex.Message, LogType.E);
                InsertErrorLog(0, ErrorType.S, "Update To Service Health Check Failed:" + ex.Message, executionServiceName, "UpdateServiceHealthCheck"); // DocumentID as 0
                return false;
            }
            return result;
        }

        private string ReadTextFileAndWriteInputOutPutInExcel(string inputvalueOriginal)
        {
            try
            {
                Match objmatch = Regex.Match(inputvalueOriginal, "[0-9]");

                if (objmatch.Success == true)
                {
                    if (inputvalueOriginal.Replace(" ", "").Trim().ToUpper().Contains("POBOX") == false)
                    {
                        string inputvalue1 = inputvalueOriginal.Substring(objmatch.Index, inputvalueOriginal.Length - objmatch.Index);
                        string[] Arrayinputvalue1WordCount = inputvalue1.Trim().Split(' ');
                        if (Regex.Match(inputvalue1, "[a-z]", RegexOptions.IgnoreCase).Success == true && inputvalue1.Trim().Split(' ').Length > 2)
                        {
                            inputvalueOriginal = inputvalue1;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
            }
            return inputvalueOriginal;
        }
        private StringBuilder AppendExtractedData_HeaderField_F3(RetStructF3 possibleWord, string FieldName, long documentID, string OBLField)
        {
            objExtractedData = new StringBuilder();

            objExtractedData.Append("<Header>");
            objExtractedData.Append("<DocumentID> " + documentID + " </DocumentID>");
            objExtractedData.Append("<FieldName> " + FieldName + " </FieldName>");

            string ContentsOfField = "";
            string ContentsOfField_Db2 = "";
            string ConfString = string.Empty;
            if (possibleWord.Words != null)
            {
                if (possibleWord.Words.Count > 0)
                {
                    if (FieldName == "Cons Name OBL" || FieldName == "Cons Addr OBL" || FieldName == "Cons City OBL" || FieldName == "Cons State OBL" || FieldName == "Cons Zip OBL")
                    {
                        ConsigneeList.Add(FieldName, possibleWord.Words[0]);
                    }

                    ContentsOfField = possibleWord.Words[0].strWord;

                    if (FieldName.Trim().ToLower() != "special_instr" && FieldName.Trim().ToLower() != "accessorial" && FieldName.Trim().ToLower() != "del_requirements")
                    {
                        ContentsOfField = Regex.Replace(ContentsOfField, "^[^a-zA-Z0-9%]+", "");
                        ContentsOfField = Regex.Replace(ContentsOfField, "[^a-zA-Z0-9)]+$", "");
                    }

                    ContentsOfField = ContentsOfField.ToUpper().Trim();
                }
            }

            if (ContentsOfField.Trim() != "")
            {
                string Remarks = "";
                string Flag = "0";
                int ConfidenceLevelofSuspect = 0;

                ContentsOfField_Db2 = ContentsOfField;
               
                ContentsOfField = ContentsOfField.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&#39;").Trim();

                if (possibleWord.ConfidenceLevelofSuspect != null)
                {
                    if (possibleWord.ConfidenceLevelofSuspect.Count > 0)
                    {
                        ConfidenceLevelofSuspect = possibleWord.ConfidenceLevelofSuspect[0];
                    }
                }


                if (possibleWord.Words[0].Remarks != null)
                {
                    Remarks = possibleWord.Words[0].Remarks;
                }
                if (possibleWord.Words[0].Flag != null)
                {
                    Flag = possibleWord.Words[0].Flag;
                }

                if (FieldName.Trim() == "Pro_No")
                {
                    Flag = "1";
                }

                if (ContentsOfField.Trim() == "*****")
                {
                    Flag = "0";
                    ConfString = "55555";
                }
                else
                {
                    ConfString = possibleWord.Words[0].ConfString;
                }



                objExtractedData.Append("<FieldValue>" + ContentsOfField + "</FieldValue>");
                objExtractedData.Append("<AvgOCRConfLevel>" + Math.Round(possibleWord.Words[0].Confidence) + "</AvgOCRConfLevel>");
                objExtractedData.Append("<CharConflevel>" + ConfString + " </CharConflevel>");
                objExtractedData.Append("<SearchConfLevel> </SearchConfLevel>");
                //objExtractedData.Append("<SearchConfLevel> 90 </SearchConfLevel>");
                objExtractedData.Append("<Remarks> " + Remarks + " </Remarks>");
                objExtractedData.Append("<PageNO>" + possibleWord.Words[0].PageNo + "</PageNO>");
                objExtractedData.Append("<ROI>" + possibleWord.Words[0].Left + ',' + possibleWord.Words[0].Top + ',' + possibleWord.Words[0].Right + ',' + possibleWord.Words[0].Bottom + "</ROI>");
                objExtractedData.Append("<FieldColor>" + Flag + "</FieldColor>");
                objExtractedData.Append("<OBLField>" + OBLField + "</OBLField>");

                if (OBLField == "1" && Flag.Contains("1"))
                {
                    SrNo += 1;
                    if (ContentsOfField_Db2.Trim() == "*****")
                    {
                        ContentsOfField_Db2 = string.Empty;
                    }


                    ContentsOfField_Db2 = ContentsOfField_Db2.Replace("'", "''");
                    dt_SortedHeaderFielddata.Rows.Add(new string[] { SrNo.ToString(), FieldName, ContentsOfField_Db2, Flag });
                }

                if (FieldName == Shp_Name_OBL || FieldName == Shp_Addr_Line_OBL || FieldName == Shp_City_OBL || FieldName == Shp_State_OBL || FieldName == Shp_ZipCode_OBL)
                {
                    if (ContentsOfField_Db2.Trim() == "*****")
                    {
                        ContentsOfField_Db2 = string.Empty;
                    }
                    ContentsOfField_Db2 = ContentsOfField_Db2.Replace("'", "''");
                    ShipperList.Add(FieldName, ContentsOfField_Db2);
                }

            }
            else
            {
                objExtractedData.Append("<FieldValue>*****</FieldValue>");
                objExtractedData.Append("<AvgOCRConfLevel>  </AvgOCRConfLevel>");
                objExtractedData.Append("<CharConflevel> </CharConflevel>");
                objExtractedData.Append("<SearchConfLevel>  </SearchConfLevel>");
                objExtractedData.Append("<Remarks>  </Remarks>");
                objExtractedData.Append("<PageNO>  </PageNO>");
                objExtractedData.Append("<ROI>  </ROI>");
                objExtractedData.Append("<FieldColor>0</FieldColor>");
                objExtractedData.Append("<OBLField>" + OBLField + "</OBLField>");
            }

            objExtractedData.Append("</Header>");

            return objExtractedData;

        }
        private StringBuilder AppendExtractedData_HeaderField_F27(RetStructF27 possibleWord, string FieldName, long documentID, string OBLField, bool isConMisc, string ConsigneeDetails_Pickup)
        {
            objExtractedData = new StringBuilder();

            objExtractedData.Append("<Header>");
            objExtractedData.Append("<DocumentID> " + documentID + " </DocumentID>");
            objExtractedData.Append("<FieldName> " + FieldName + " </FieldName>");


            string ContentsOfField = "";
            string ContentsOfField_Db2 = "";
            string Flag = "0";

            if (isConMisc == true)
            {
                if (FieldName.Trim() == "Con_Code")
                {
                    if (CheckZipORCityState_MatchWithConsPickUp(ConsigneeDetails_Pickup) == true)
                    {
                        possibleWord.ContentsOfField = "000000N";
                        possibleWord.Flag = "1";
                    }
                }

                if (FieldName.Trim() == "Con_Name")
                {
                    if (CheckZipORCityState_MatchWithConsPickUp(ConsigneeDetails_Pickup) == true)
                    {
                        if (ConsigneeList.ContainsKey("Cons Name OBL") == true)
                        {
                            Regex regNotStartWithNumber = new Regex("^\\d[^<]+"); // Added on 12-Aug-2021 => OBL Name should not start with numeric
                            string conName = ConsigneeList["Cons Name OBL"].strWord;
                            conName = Regex.Replace(conName, "^[^a-zA-Z0-9%]+", "");
                            conName = Regex.Replace(conName, "[^a-zA-Z0-9)]+$", "");

                            if (ConsigneeList["Cons Name OBL"].Confidence >= 85 && regNotStartWithNumber.IsMatch(conName) == false && conName.Trim() != "")
                            {
                                possibleWord.ContentsOfField = conName;
                                possibleWord.Flag = "1";
                            }
                        }


                    }
                }

                if (FieldName.Trim() == "Con_Addr_Line")
                {
                    if (CheckZipORCityState_MatchWithConsPickUp(ConsigneeDetails_Pickup) == true)
                    {
                        if (ConsigneeList.ContainsKey("Cons Addr OBL") == true)
                        {
                            string conAddres = ConsigneeList["Cons Addr OBL"].strWord;
                            conAddres = Regex.Replace(conAddres, "^[^a-zA-Z0-9%]+", "");
                            conAddres = Regex.Replace(conAddres, "[^a-zA-Z0-9)]+$", "");

                            conAddres = ReadTextFileAndWriteInputOutPutInExcel(conAddres);

                            if (ConsigneeList["Cons Addr OBL"].Confidence >= 85 && conAddres.Trim() != "")
                            {
                                possibleWord.ContentsOfField = conAddres;
                                possibleWord.Flag = "1";
                            }
                        }
                    }
                }
            }

            if (possibleWord.ContentsOfField != null)
            {

                ContentsOfField = possibleWord.ContentsOfField;

                if (FieldName.ToLower() != "special_instr" && FieldName.ToLower() != "accessorial" && FieldName.ToLower() != "del_requirements")
                {
                    ContentsOfField = Regex.Replace(ContentsOfField, "^[^a-zA-Z0-9%]+", "");
                    ContentsOfField = Regex.Replace(ContentsOfField, "[^a-zA-Z0-9)]+$", "");
                }

                ContentsOfField = ContentsOfField.ToUpper().Trim();
            }

            if (ContentsOfField.Trim() != "")
            {
                string Remarks = "";
                ContentsOfField_Db2 = ContentsOfField;

                ContentsOfField = ContentsOfField.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&#39;").Trim();

                if (possibleWord.Remarks != null)
                {
                    Remarks = possibleWord.Remarks;
                }
                if (possibleWord.Flag != null)
                {
                    Flag = possibleWord.Flag;
                }

                if (ContentsOfField.Trim() == "*****")
                {
                    Flag = "0";
                }



                objExtractedData.Append("<FieldValue>" + ContentsOfField + "</FieldValue>");
                objExtractedData.Append("<AvgOCRConfLevel>" + Math.Round(Convert.ToDouble(possibleWord.ConfLevel)) + "</AvgOCRConfLevel>");
                objExtractedData.Append("<CharConflevel>" + possibleWord.ConfString + " </CharConflevel>");
                //objExtractedData.Append("<SearchConfLevel>" + possibleWord.ConfidenceLevelofSuspect[0] + "</SearchConfLevel>");
                objExtractedData.Append("<SearchConfLevel>  </SearchConfLevel>");
                objExtractedData.Append("<Remarks> " + Remarks + " </Remarks>");
                objExtractedData.Append("<PageNO>" + possibleWord.PageNo + "</PageNO>");
                objExtractedData.Append("<ROI>" + possibleWord.X1 + ',' + possibleWord.Y1 + ',' + possibleWord.X2 + ',' + possibleWord.Y2 + "</ROI>");
                objExtractedData.Append("<FieldColor>" + Flag + "</FieldColor>");
                objExtractedData.Append("<OBLField>" + OBLField + "</OBLField>");

                if (FieldName.Trim() == "Con_Code")
                {
                    Cons_Code_Misc = "<FieldValue>" + ContentsOfField + "</FieldValue>";
                }
            }
            else
            {
                objExtractedData.Append("<FieldValue>*****</FieldValue>");
                objExtractedData.Append("<AvgOCRConfLevel>  </AvgOCRConfLevel>");
                objExtractedData.Append("<CharConflevel> </CharConflevel>");
                objExtractedData.Append("<SearchConfLevel>  </SearchConfLevel>");
                objExtractedData.Append("<Remarks>  </Remarks>");
                objExtractedData.Append("<PageNO>  </PageNO>");
                objExtractedData.Append("<ROI>  </ROI>");
                objExtractedData.Append("<FieldColor>0</FieldColor>");
                objExtractedData.Append("<OBLField>" + OBLField + "</OBLField>");

                if (FieldName.Trim() == "Con_Code")
                {
                    Cons_Code_Misc = "<FieldValue>*****</FieldValue>";
                }
            }

            if (OBLField == "1")
            {

                if (ContentsOfField_Db2.Trim() == "*****")
                {
                    ContentsOfField_Db2 = "";
                }

                ContentsOfField_Db2 = ContentsOfField_Db2.Replace("'", "''");
                
                if (FieldName == Shp_Code || FieldName == Shp_Name || FieldName == Shp_Addr_Line || FieldName == Shp_City || FieldName == Shp_State || FieldName == Shp_ZipCode)
                {
                    if (Flag.Contains("1") == false)
                    {
                        Shp_ChangeGreenToRed = true;
                    }
                    SrNo += 1;
                    dt_SortedHeaderFielddata.Rows.Add(new string[] { SrNo.ToString(), FieldName, ContentsOfField_Db2, Flag });

                    ShipperList.Add(FieldName, ContentsOfField_Db2);
                }
                else if (FieldName == Con_Code || FieldName == Con_Name || FieldName == Con_Addr_Line || FieldName == Con_City || FieldName == Con_State || FieldName == Con_ZipCode)
                {
                    if (Flag.Contains("1") == false && FieldName != Con_Code)
                    {
                        Con_ChangeGreenToRed = true;
                    }

                    SrNo += 1;
                    dt_SortedHeaderFielddata.Rows.Add(new string[] { SrNo.ToString(), FieldName, ContentsOfField_Db2, Flag });
                }
                else if (FieldName == Bill_To_Code || FieldName == Bill_To_Name || FieldName == Bill_To_Addr_Line || FieldName == Bill_To_City || FieldName == Bill_To_State || FieldName == Bill_To_ZipCode)
                {
                    if (Flag.Contains("1") == false)
                    {
                        BillTo_ChangeGreenToRed = true;
                    }

                    SrNo += 1;
                    dt_SortedHeaderFielddata.Rows.Add(new string[] { SrNo.ToString(), FieldName, ContentsOfField_Db2, Flag });
                }
                else if (Flag.Contains("1") == true)
                {
                    SrNo += 1;
                    dt_SortedHeaderFielddata.Rows.Add(new string[] { SrNo.ToString(), FieldName, ContentsOfField_Db2, Flag });
                }
            }


            if (FieldName == Shp_Code_PKP || FieldName == ShipperDetails_PKP)
            {
                if (ContentsOfField_Db2.Trim() == "*****")
                {
                    ContentsOfField_Db2 = "";
                }

                ContentsOfField_Db2 = ContentsOfField_Db2.Replace("'", "''");
                ShipperList.Add(FieldName, ContentsOfField_Db2);
            }

            objExtractedData.Append("</Header>");

            return objExtractedData;

        }
        private StringBuilder AppendExtractedData_LineField(clsCnCWord possibleWord, string FieldName, long documentID, int RowNo, bool iswordFound, string OBLField)
        {
            objExtractedData = new StringBuilder();

            objExtractedData.Append("<Row>");
            objExtractedData.Append("<DocumentID>" + documentID + "</DocumentID>");
            objExtractedData.Append("<FieldName>" + FieldName + "</FieldName>");


            string ContentsOfField = "";
            string ContentsOfField_Db2 = "";
            string Flag = "0";
            string ConfString = string.Empty;

            if (iswordFound == true)
            {
                if (possibleWord.strWord != null)
                {

                    ContentsOfField = possibleWord.strWord;
                    ContentsOfField = Regex.Replace(ContentsOfField, "^[^a-zA-Z0-9%]+", "");
                    ContentsOfField = Regex.Replace(ContentsOfField, "[^a-zA-Z0-9)]+$", "");
                    ContentsOfField = ContentsOfField.ToUpper().Trim();

                }
            }


            if (iswordFound == true && ContentsOfField.Trim() != "")
            {
                string Remarks = "";
                ContentsOfField_Db2 = ContentsOfField;
                
                ContentsOfField = ContentsOfField.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&#39;").Trim();

                if (possibleWord.Remarks != null)
                {
                    Remarks = possibleWord.Remarks;
                }
                if (possibleWord.Flag != null)
                {
                    Flag = possibleWord.Flag;
                }

                if (ContentsOfField.Trim() == "*****")
                {
                    Flag = "0";
                    ConfString = "55555";
                }
                else
                {
                    ConfString = possibleWord.ConfString;
                }


                objExtractedData.Append("<FieldValue>" + ContentsOfField + "</FieldValue>");
                objExtractedData.Append("<AvgOCRConfLevel>" + Math.Round(possibleWord.Confidence) + "</AvgOCRConfLevel>");
                objExtractedData.Append("<CharConflevel>" + ConfString + " </CharConflevel>");
                // objExtractedData.Append("<SearchConfLevel>" + possibleWord.ConfidenceLevelofSuspect[0] + "</SearchConfLevel>");
                objExtractedData.Append("<SearchConfLevel>  </SearchConfLevel>");
                objExtractedData.Append("<Remarks> " + Remarks + " </Remarks>");
                objExtractedData.Append("<PageNO>" + possibleWord.PageNo + "</PageNO>");
                objExtractedData.Append("<RowNo>" + RowNo + "</RowNo>");
                objExtractedData.Append("<ROI>" + possibleWord.Left + ',' + possibleWord.Top + ',' + possibleWord.Right + ',' + possibleWord.Bottom + "</ROI>");
                objExtractedData.Append("<FieldColor>" + Flag + "</FieldColor>");
                objExtractedData.Append("<OBLField>" + OBLField + "</OBLField>");
            }
            else
            {
                objExtractedData.Append("<FieldValue>*****</FieldValue>");
                objExtractedData.Append("<AvgOCRConfLevel>  </AvgOCRConfLevel>");
                objExtractedData.Append("<CharConflevel> </CharConflevel>");
                objExtractedData.Append("<SearchConfLevel>  </SearchConfLevel>");
                objExtractedData.Append("<Remarks>  </Remarks>");
                objExtractedData.Append("<PageNO>  </PageNO>");
                objExtractedData.Append("<RowNo>" + RowNo + "</RowNo>");
                objExtractedData.Append("<ROI>  </ROI>");
                objExtractedData.Append("<FieldColor>" + Flag + "</FieldColor>");
                objExtractedData.Append("<OBLField>" + OBLField + "</OBLField>");
            }


            if (OBLField == "1")
            {
                SrNo += 1;

                if (ContentsOfField_Db2.Trim() == "*****")
                {
                    ContentsOfField_Db2 = "";
                }

                ContentsOfField_Db2 = ContentsOfField_Db2.Replace("'", "''");

                dt_SortedLineitemdata.Rows.Add(new string[] { SrNo.ToString(), FieldName, ContentsOfField_Db2, Flag, RowNo.ToString() });
            }


            objExtractedData.Append("</Row>");

            return objExtractedData;

        }

        public bool CheckZipORCityState_MatchWithConsPickUp(string ConsigneeDetails_Pickup)
        {
            bool result = false;
            try
            {
                string[] arry_shipperDetails_PKP = ConsigneeDetails_Pickup.Split('~');
                if (arry_shipperDetails_PKP.Length == 5)
                {
                    string ConZip_OBL = "";
                    string ConCity_OBL = "";
                    string ConState_OBL = "";

                    if (ConsigneeList.ContainsKey("Cons Zip OBL") == true)
                    {
                        ConZip_OBL = ConsigneeList["Cons Zip OBL"].strWord;
                    }
                    if (ConsigneeList.ContainsKey("Cons City OBL") == true)
                    {
                        ConCity_OBL = ConsigneeList["Cons City OBL"].strWord;
                    }
                    if (ConsigneeList.ContainsKey("Cons State OBL") == true)

                    {
                        ConState_OBL = ConsigneeList["Cons State OBL"].strWord;
                    }

                    ConZip_OBL = Regex.Replace(ConZip_OBL, "^[^a-zA-Z0-9%]+", "");
                    ConZip_OBL = Regex.Replace(ConZip_OBL, "[^a-zA-Z0-9)]+$", "");

                    ConCity_OBL = Regex.Replace(ConCity_OBL, "^[^a-zA-Z0-9%]+", "");
                    ConCity_OBL = Regex.Replace(ConCity_OBL, "[^a-zA-Z0-9)]+$", "");

                    ConState_OBL = Regex.Replace(ConState_OBL, "^[^a-zA-Z0-9%]+", "");
                    ConState_OBL = Regex.Replace(ConState_OBL, "[^a-zA-Z0-9)]+$", "");

                    if ((arry_shipperDetails_PKP[4].Trim() == ConZip_OBL.Trim() && ConZip_OBL.Trim() != "") || (arry_shipperDetails_PKP[2].Trim().Replace(" ", "").ToUpper() == ConCity_OBL.Trim().Replace(" ", "").ToUpper() && arry_shipperDetails_PKP[3].Trim().Replace(" ", "").ToUpper() == ConState_OBL.Trim().Replace(" ", "").ToUpper() && ConState_OBL.Trim() != ""))
                    {
                        if (ConsigneeList.ContainsKey("Cons Name OBL") == true)
                        {
                            Regex regNotStartWithNumber = new Regex("^\\d[^<]+"); // Added on 12-Aug-2021 => OBL Name should not start with numeric
                            string conName = ConsigneeList["Cons Name OBL"].strWord;
                            conName = Regex.Replace(conName, "^[^a-zA-Z0-9%]+", "");
                            conName = Regex.Replace(conName, "[^a-zA-Z0-9)]+$", "");

                            if (ConsigneeList["Cons Name OBL"].Confidence >= 85 && regNotStartWithNumber.IsMatch(conName) == false && conName.Trim() != "")
                            {
                                result = true;
                            }
                        }
                        if (ConsigneeList.ContainsKey("Cons Addr OBL") == true)
                        {
                            string conAddres = ConsigneeList["Cons Addr OBL"].strWord;
                            conAddres = Regex.Replace(conAddres, "^[^a-zA-Z0-9%]+", "");
                            conAddres = Regex.Replace(conAddres, "[^a-zA-Z0-9)]+$", "");

                            if (ConsigneeList["Cons Addr OBL"].Confidence >= 85 && conAddres.Trim() != "")
                            {
                                result = true;
                            }
                        }
                    }
                }

            }
            catch (Exception e)
            {
                result = false;
            }

            return result;
        }

        public bool InsertShipperDetails_KM2(long documentID, string tiffName)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            bool result = false;
            try
            {

                if (!string.IsNullOrEmpty(DBUtilities.prodConnectionString) && documentID > 0)
                {
                    string spName = "USP_ShipperKM_Insert";

                    SqlParameter documentIDParameter = new SqlParameter("@DocumentID", documentID);
                    SqlParameter tiffnameParameter = new SqlParameter("@Tiffname", tiffName);


                    SqlParameter shipCodeParameter = new SqlParameter("@ShipCode", ShipperList["Shp_Code"]);
                    SqlParameter shipName_FNLParameter = new SqlParameter("@ShipName_FNL", ShipperList["Shp_Name"]);
                    SqlParameter shipAddress_FNLParameter = new SqlParameter("@ShipAddress_FNL", ShipperList["Shp_Addr_Line"]);
                    SqlParameter shipCity_FNLParameter = new SqlParameter("@ShipCity_FNL", ShipperList["Shp_City"]);
                    SqlParameter shipState_FNLParameter = new SqlParameter("@ShipState_FNL", ShipperList["Shp_State"]);
                    SqlParameter ship_Zip_FNLParameter = new SqlParameter("@Ship_Zip_FNL", ShipperList["Shp_ZipCode"]);

                    SqlParameter shipName_OBLParameter = new SqlParameter("@ShipName_OBL", ShipperList["Ship Name OBL"]);
                    SqlParameter shipAddress_OBLParameter = new SqlParameter("@ShipAddress_OBL", ShipperList["Ship Addr OBL"]);
                    SqlParameter shipCity_OBLParameter = new SqlParameter("@ShipCity_OBL", ShipperList["Ship City OBL"]);
                    SqlParameter shipState_OBLParameter = new SqlParameter("@ShipState_OBL", ShipperList["Ship State OBL"]);
                    SqlParameter ship_Zip_OBLParameter = new SqlParameter("@Ship_Zip_OBL", ShipperList["Ship Zip OBL"]);

                    SqlParameter shipCodePKP_Parameter = new SqlParameter("@ShipCode_PKP", ShipperList["ShipperCode_Pickup"]);

                    string ShipperDetails_Pickup = ShipperList["ShipperDetails_Pickup"];
                    string[] arry_shipperDetails_PKP = ShipperDetails_Pickup.Split('~');
                    string shipNamePKPParameter = "";
                    string shipAddressPKPParameter = "";
                    string shipCityPKPParameter = "";
                    string shipStatePKPParameter = "";
                    string shipZipCodePKPParameter = "";
                    if (arry_shipperDetails_PKP.Length > 0)
                    {
                        shipNamePKPParameter = arry_shipperDetails_PKP[0];
                    }
                    if (arry_shipperDetails_PKP.Length > 1)
                    {
                        shipAddressPKPParameter = arry_shipperDetails_PKP[1];
                    }
                    if (arry_shipperDetails_PKP.Length > 2)
                    {
                        shipCityPKPParameter = arry_shipperDetails_PKP[2];
                    }
                    if (arry_shipperDetails_PKP.Length > 3)
                    {
                        shipStatePKPParameter = arry_shipperDetails_PKP[3];
                    }
                    if (arry_shipperDetails_PKP.Length > 4)
                    {
                        shipZipCodePKPParameter = arry_shipperDetails_PKP[4];
                    }


                    SqlParameter shipNamePKP_Parameter = new SqlParameter("@ShipName_PKP", shipNamePKPParameter);
                    SqlParameter shipAddressPKP_Parameter = new SqlParameter("@ShipAddress_PKP", shipAddressPKPParameter);
                    SqlParameter shipCityPKP_Parameter = new SqlParameter("@ShipCity_PKP", shipCityPKPParameter);
                    SqlParameter shipStatePKP_Parameter = new SqlParameter("@ShipState_PKP", shipStatePKPParameter);
                    SqlParameter shipZipCodePKP_Parameter = new SqlParameter("@Ship_Zip_PKP", shipZipCodePKPParameter);

                    SqlParameter shipperDetailsPKP_Parameter = new SqlParameter("@ShipperDetails_PKP", ShipperDetails_Pickup);

                    SqlParameter[] parameters = new SqlParameter[] { documentIDParameter, tiffnameParameter, shipCodeParameter, shipName_FNLParameter, shipAddress_FNLParameter, shipCity_FNLParameter, shipState_FNLParameter, ship_Zip_FNLParameter, shipName_OBLParameter, shipAddress_OBLParameter, shipCity_OBLParameter, shipState_OBLParameter, ship_Zip_OBLParameter, shipCodePKP_Parameter, shipNamePKP_Parameter, shipAddressPKP_Parameter, shipCityPKP_Parameter, shipStatePKP_Parameter, shipZipCodePKP_Parameter, shipperDetailsPKP_Parameter };
                    serviceResponse = DBUtilities.InsertOrUpdateBySP(DBUtilities.prodConnectionString, spName, parameters);
                    if (serviceResponse.Status == ServiceResponseStatus.Success)
                    {
                        result = true;
                    }
                }
                else
                {
                    WriteLog(documentID + ": " + "Insert to InsertShipperDetails_KM2 failed due to  : prodConnectionString is empty or document id  0", LogType.P);
                }

            }
            catch (Exception ex)
            {
                WriteLog(documentID + ": " + "Insert to InsertShipperDetails_KM2 failed:" + ex.Message, LogType.E);
                return false;
            }
            return result;
        }
        private bool GetEDIStatus(string str_ProNumber)
        {
            bool isEDI = false;

            try
            {
                

                string queryString = "SELECT * from FRP001ED where FHPRO in (SELECT PCEPRO FROM DSP092 where PCPRO in ('" + str_ProNumber + "'))";

                isEDI = DBUtilities.GetStatusFromFRP001(queryString);

                //return isEDI;
            }
            catch (Exception ex)
            {
                isEDI = false;
                WriteLog("GetEDIStatus: " + "select EDI Status from FRP001ED and DSP092 for PRO " + str_ProNumber + "  failed: " + ex.Message, LogType.E);
            }

            return isEDI;
        }

        private bool GetBillingStatus(string str_ProNumber)
        {
            bool Billed = false;

            try
            {
                //select FHPRO, FHSTAT from FRP001 where FHPRO = 77005880760 and FHSTAT = 'PP'

                string queryString = "select FHPRO, FHSTAT from FRP001 where FHPRO = " + str_ProNumber + " and FHSTAT = '" + billingStatus + "' ";

                Billed = DBUtilities.GetStatusFromFRP001(queryString);

            }
            catch (Exception ex)
            {
                Billed = false;
                WriteLog("GetBillingStatus: " + "select FHSTAT status from FRP001 for PRO " + str_ProNumber + "  failed: " + ex.Message, LogType.E);
            }

            return Billed;
        }


        public static DataTable GetDocumentMasterData()
        {
            ServiceResponse serviceResponse = new ServiceResponse();

            try
            {
                if (!string.IsNullOrEmpty(DBUtilities.prodConnectionString))
                {
                    string instanceId = ConfigurationManager.AppSettings["InstanceId"].ToString();
                    string spName = "USP_DocumentMaster_GetdataForExtraction";
                    SqlParameter serverNameParameter = new SqlParameter("@ServerName", executionServerName);
                    SqlParameter currentStatusParameter = new SqlParameter("@CurrentStatus", DocumentStatus.Metadata); //1 is status which can be Downloaded, Preprocessed, Metadata etc.
                    SqlParameter serviceNameParameter = new SqlParameter("@ServiceName", executionServiceName);
                    SqlParameter instanceIdParameter = new SqlParameter("@InstanceId", Convert.ToInt32(instanceId));
                    SqlParameter[] parameters = new SqlParameter[] { serverNameParameter, currentStatusParameter, serviceNameParameter, instanceIdParameter };
                    serviceResponse = DBUtilities.GetDataTableFromSP(DBUtilities.prodConnectionString, spName, parameters);
                }
            }
            catch (Exception ex)
            {
                WriteLog("GetDocumentMasterData: " + "select from documentmaster failed:" + ex.Message, LogType.E);

            }
            return serviceResponse.Response;
        }

        public static  bool  UpdateDocumentStatus_As400Color(long documentID, int CurrentStatus, string Remarks, DateTime StartDt, DateTime EndDt, string Username)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            bool result = false;
            try
            {
                if (!string.IsNullOrEmpty(DBUtilities.prodConnectionString))
                {
                    string spName = "USP_Update_DocumentStatus";

                    SqlParameter DocIdParam = new SqlParameter("@DocumentID", documentID);
                    SqlParameter CurrStatusParam = new SqlParameter("@CurrentStatus", CurrentStatus);
                    SqlParameter RemarksParam = new SqlParameter("@Remarks", Remarks);
                    SqlParameter StartDtParam = new SqlParameter("@StartDate", StartDt);
                    SqlParameter EndDtParam = new SqlParameter("@EndDate", EndDt);
                    SqlParameter UserNameParam = new SqlParameter("@Username", Username);
                    SqlParameter[] parameters = new SqlParameter[] { DocIdParam, CurrStatusParam, RemarksParam, StartDtParam, EndDtParam, UserNameParam };
                    serviceResponse = DBUtilities.InsertOrUpdateBySP(DBUtilities.prodConnectionString, spName, parameters);
                    if (serviceResponse.Status == ServiceResponseStatus.Success)
                    {
                        result = true;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog(documentID + ": " + " Update to UpdateDocumentStatus_As400Color failed:" + ex.Message, LogType.E);
                InsertErrorLog(documentID, ErrorType.S, documentID + ":" + " Update to UpdateDocumentStatus_As400Color failed:" + ex.Message, executionServiceName, "UpdateDocumentStatus_As400Color");
                return false;
               
            }
            return result;
        }
        public static bool UpdateDocumentMaster(long documentID, DateTime startTime, DateTime endTime, string TotaltimeInSec, int imgOrentation, bool isEDIDoc)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            bool result = false;
            DateTime startTimeupdate = DateTime.Now;
            try
            {
               // WriteLog("Start to update UpdateDocumentMaster() on DB", LogType.P);

                if (!string.IsNullOrEmpty(DBUtilities.prodConnectionString) && documentID > 0)
                {
                    string spName = "USP_DocumentMaster_Update";
                    SqlParameter documentIDParameter = new SqlParameter("@DocumentID", documentID);
                    SqlParameter currentStatusParameter = new SqlParameter("@CurrentStatus", DocumentStatus.Extracted); //4 is status which can be Downloaded, Preprocessed, Metadata etc.
                    SqlParameter remarksParameter = new SqlParameter("@Remarks", "DataExtracted-" + TotaltimeInSec);
                    SqlParameter startDateParameter = new SqlParameter("@StartDate", startTime.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                    SqlParameter endDateParameter = new SqlParameter("@EndDate", endTime.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                    SqlParameter serviceNameParameter = new SqlParameter("@ServiceName", executionServiceName);
                    SqlParameter serverNameParameter = new SqlParameter("@ServerName", executionServerName);
                    SqlParameter imgOrentationParameter = new SqlParameter("@RotationAngle", imgOrentation);
                    SqlParameter isEDIDocParameter = new SqlParameter("@IsEdi", isEDIDoc);
                    SqlParameter[] parameters = new SqlParameter[] { documentIDParameter, currentStatusParameter, remarksParameter, startDateParameter, endDateParameter, serviceNameParameter, serverNameParameter, imgOrentationParameter, isEDIDocParameter };
                    serviceResponse = DBUtilities.InsertOrUpdateBySP(DBUtilities.prodConnectionString, spName, parameters);
                    if (serviceResponse.Status == ServiceResponseStatus.Success)
                    {
                        result = true;
                    }
                    else
                    {
                        WriteLog(documentID + ": " + " Update to documentmaster failed:" + serviceResponse.Message , LogType.E);
                        InsertErrorLog(documentID, ErrorType.S, documentID + ":" + " Update to documentmaster failed:" + serviceResponse.Message, executionServiceName, "UpdateDocumentMaster");
                    }

                }
                else
                {
                    WriteLog("UpdateDocumentMaster failed: Dueto: " + documentID + ": " + " prodConnectionString: " + DBUtilities.prodConnectionString, LogType.E);
                }

               // WriteLog("End to update UpdateDocumentMaster() on DB", LogType.P);

                DateTime startTime_Lineupdate = DateTime.Now;
                WriteLog("Total Time Taken for UpdateDocumentMaster(): " + CommonModule.Common.CalculateTotalTime(startTimeupdate, startTime_Lineupdate), LogType.P);



            }
            catch (Exception ex)
            {
                WriteLog(documentID + ": " + " Update to documentmaster failed:" + ex.Message, LogType.E);
                InsertErrorLog(documentID, ErrorType.S, documentID + ":" + " Update to documentmaster failed:" + ex.Message, executionServiceName, "UpdateDocumentMaster");
                return false;
            }
            return result;
        }
        public static bool UpdateDocumentMaster_ErrorFlag(long documentID, DateTime startTime, DateTime endTime)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            bool result = false;
            try
            {

                if (!string.IsNullOrEmpty(DBUtilities.prodConnectionString) && documentID > 0)
                {
                    string spName = "USP_DocumentMaster_Update";
                    SqlParameter documentIDParameter = new SqlParameter("@DocumentID", documentID);
                    SqlParameter currentStatusParameter = new SqlParameter("@CurrentStatus", DocumentStatus.Failed); //5 is status for Failed.
                    SqlParameter remarksParameter = new SqlParameter("@Remarks", "DataExtractionFailed");
                    SqlParameter startDateParameter = new SqlParameter("@StartDate", startTime.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                    SqlParameter endDateParameter = new SqlParameter("@EndDate", endTime.ToString("yyyy-MM-dd HH:mm:ss.fff"));

                    SqlParameter serviceNameParameter = new SqlParameter("@ServiceName", executionServiceName);
                    SqlParameter serverNameParameter = new SqlParameter("@ServerName", executionServerName);

                    SqlParameter imgOrentationParameter = new SqlParameter("@RotationAngle", 0);

                    SqlParameter[] parameters = new SqlParameter[] { documentIDParameter, currentStatusParameter, remarksParameter, startDateParameter, endDateParameter, serviceNameParameter, serverNameParameter, imgOrentationParameter };
                    serviceResponse = DBUtilities.InsertOrUpdateBySP(DBUtilities.prodConnectionString, spName, parameters);
                    if (serviceResponse.Status == ServiceResponseStatus.Success)
                    {
                        result = true;
                    }
                }
                else
                {
                    WriteLog(documentID + ": " + "Update to UpdateDocumentMaster_ErrorFlag failed not updated due to  : prodConnectionString is empty or document id  0", LogType.P);
                }
                
            }
            catch (Exception ex)
            {
                WriteLog(documentID + ": " + "Update to UpdateDocumentMaster_ErrorFlag failed:" + ex.Message, LogType.E);
                //InsertErrorLog(documentID, ErrorType.S, documentID + ":" + "Update to documentmaster failed:" + ex.Message, executionServiceName, "UpdateDocumentMaster_ErrorFlag");
                return false;
            }
            return result;
        }

        public static bool HeaderFieldData_BulkInsert(long documentID, string DataExtraction_HeaderFieldData)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            bool result = false;

            DateTime startTime = DateTime.Now;


            try
            {
               // WriteLog(" Start to Insert HeaderFieldData() on DB ", LogType.P);
                if (!string.IsNullOrEmpty(DataExtraction_HeaderFieldData))
                {
                    string spName = "USP_HeaderFieldData_BulkInsert";
                    SqlParameter headerFieldDataParameter = new SqlParameter("@HeaderDetails", DataExtraction_HeaderFieldData);


                    SqlParameter[] parameters = new SqlParameter[] { headerFieldDataParameter };
                    serviceResponse = DBUtilities.InsertOrUpdateBySP(DBUtilities.prodConnectionString, spName, parameters);
                    if (serviceResponse.Status == ServiceResponseStatus.Success)
                    {
                        result = true;
                    }
                    else
                    {
                        WriteLog(documentID + ": " + "Insert to headerfielddata failed:" + serviceResponse.Message, LogType.E);
                        InsertErrorLog(documentID, ErrorType.S, documentID + ":" + "Insert to headerfielddata failed:" + serviceResponse.Message, executionServiceName, "HeaderFieldData_BulkInsert");
                    }

                   // WriteLog("End to Insert HeaderFieldData() on DB", LogType.P);

                    DateTime startTime_Line = DateTime.Now;

                    WriteLog("Total Time Taken for HeaderFieldData(): " + CommonModule.Common.CalculateTotalTime(startTime, startTime_Line), LogType.P);



                }
            }
            catch (Exception ex)
            {
                WriteLog(documentID + ": " + "Insert to headerfielddata failed:" + ex.Message, LogType.E);
                InsertErrorLog(documentID, ErrorType.S, documentID + ":" + "Insert to headerfielddata failed:" + ex.Message, executionServiceName, "HeaderFieldData_BulkInsert");
                //UpdateDocumentMaster_ErrorFlag(documentID, DateTime.Now, DateTime.Now);
                return false;
            }
            return result;
        }
        public static bool LineFieldData_BulkInsert(long documentID, string DataExtraction_LineFieldData)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            bool result = false;

            DateTime startTime = DateTime.Now;
            

            try
            {
               // WriteLog("Start to Insert LineFieldData() on DB", LogType.P);
                if (!string.IsNullOrEmpty(DataExtraction_LineFieldData))
                {
                    string spName = "USP_LineFieldData_BulkInsert";
                    SqlParameter lineFieldDataParameter = new SqlParameter("@LineDetails", DataExtraction_LineFieldData);


                    SqlParameter[] parameters = new SqlParameter[] { lineFieldDataParameter };
                    serviceResponse = DBUtilities.InsertOrUpdateBySP(DBUtilities.prodConnectionString, spName, parameters);

                    if (serviceResponse.Status == ServiceResponseStatus.Success)
                    {
                        result = true;
                    }
                    else
                    {
                        WriteLog(documentID + ": " + "Insert to linefielddata failed:" + serviceResponse.Message, LogType.E);
                        InsertErrorLog(documentID, ErrorType.S, documentID + ":" + "Insert to linefielddata failed:" + serviceResponse.Message, executionServiceName, "LineFieldData_BulkInsert");
                    }

                    //WriteLog("End to Insert LineFieldData() on DB", LogType.P);
                    DateTime startTime_Line = DateTime.Now;

                    WriteLog("Total Time Taken for LineFieldData(): " + CommonModule.Common.CalculateTotalTime(startTime, startTime_Line), LogType.P);
                }
            }
            catch (Exception ex)
            {
                WriteLog(documentID + ": " + "Insert to linefielddata failed:" + ex.Message, LogType.E);
                InsertErrorLog(documentID, ErrorType.S, documentID + ":" + "Insert to linefielddata failed:" + ex.Message, executionServiceName, "LineFieldData_BulkInsert");
                //UpdateDocumentMaster_ErrorFlag(documentID, DateTime.Now, DateTime.Now);
                return false;
            }
            return result;
        }

        public static void InsertErrorLog(long documentID, ErrorType errorType, string errorDetails, string moduleName, string functionName)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            try
            {
                if (!string.IsNullOrEmpty(DBUtilities.prodConnectionString) && !string.IsNullOrEmpty(errorType.ToString()) && !string.IsNullOrEmpty(errorDetails) && !string.IsNullOrEmpty(moduleName) && !string.IsNullOrEmpty(functionName))
                {
                    string spName = "USP_Errorlog_Insert";
                    SqlParameter documentIDParameter = new SqlParameter("@DocumentID", documentID);
                    SqlParameter errorTypeParameter = new SqlParameter("@ErrorType", errorType.ToString());
                    SqlParameter errorDetailsParameter = new SqlParameter("@ErrorDetails", errorDetails);
                    SqlParameter moduleNameParameter = new SqlParameter("@ModuleName", moduleName);
                    SqlParameter functionNameParameter = new SqlParameter("@FunctionName", functionName);
                    SqlParameter serverNameParameter = new SqlParameter("@ServerName", executionServerName); // server on which service is running
                    SqlParameter[] parameters = new SqlParameter[] { documentIDParameter, errorTypeParameter, errorDetailsParameter, moduleNameParameter, functionNameParameter, serverNameParameter };
                    serviceResponse = DBUtilities.GetDataTableFromSP(DBUtilities.prodConnectionString, spName, parameters);
                }

                UpdateDocumentMaster_ErrorFlag(documentID, DateTime.Now, DateTime.Now);


            }
            catch (Exception ex)
            {
                WriteLog(documentID + ": " + "Insert to errorlog failed:" + ex.Message, LogType.E);
            }
        }

        public static void InsertErrorLog_As400(long documentID, ErrorType errorType, string errorDetails, string moduleName, string functionName)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            try
            {
                if (!string.IsNullOrEmpty(DBUtilities.prodConnectionString) && !string.IsNullOrEmpty(errorType.ToString()) && !string.IsNullOrEmpty(errorDetails) && !string.IsNullOrEmpty(moduleName) && !string.IsNullOrEmpty(functionName))
                {
                    string spName = "USP_Errorlog_Insert";
                    SqlParameter documentIDParameter = new SqlParameter("@DocumentID", documentID);
                    SqlParameter errorTypeParameter = new SqlParameter("@ErrorType", errorType.ToString());
                    SqlParameter errorDetailsParameter = new SqlParameter("@ErrorDetails", errorDetails);
                    SqlParameter moduleNameParameter = new SqlParameter("@ModuleName", moduleName);
                    SqlParameter functionNameParameter = new SqlParameter("@FunctionName", functionName);
                    SqlParameter serverNameParameter = new SqlParameter("@ServerName", executionServerName); // server on which service is running
                    SqlParameter[] parameters = new SqlParameter[] { documentIDParameter, errorTypeParameter, errorDetailsParameter, moduleNameParameter, functionNameParameter, serverNameParameter };
                    serviceResponse = DBUtilities.GetDataTableFromSP(DBUtilities.prodConnectionString, spName, parameters);
                }

               // UpdateDocumentMaster_ErrorFlag(documentID, DateTime.Now, DateTime.Now);


            }
            catch (Exception ex)
            {
                WriteLog(documentID + ": " + "Insert to errorlog failed:" + ex.Message, LogType.E);
            }
        }

        public static void WriteLog(string message, LogType logType)
        {
            try
            {
                if (processLogWriter != null)
                {
                    if (logType == LogType.P)
                    {
                        processLogWriter.WriteLine(DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss") + " - " + message);
                        processLogWriter.Flush();
                    }
                    else if (logType == LogType.E)
                    {
                        processLogWriter.WriteLine("----------------Error Occured----------------");
                        processLogWriter.WriteLine(DateTime.Now.ToString("MM-dd-yyyy HH:mm:ss") + " - " + message);
                        processLogWriter.Flush();
                    }
                }
            }
            catch
            {

            }
        }

    }
}
