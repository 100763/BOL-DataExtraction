using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BOLDocumentDataExtraction.CommonModule;
using System.Windows.Forms;

namespace BOLDocumentDataExtraction.DAL
{
    public static class DBUtilities
    {
        #region "Variable Decl"
        public static string appConnectionString;
        public static string prodConnectionString;
        public static string DB2ConnectionString;
        public static OdbcConnection db2Connection = new OdbcConnection();
        public static SqlConnection sqlConnection = new SqlConnection();
        #endregion

        public static bool  GetStatusFromFRP001( string queryString,OdbcParameter[] parameters = null)
        {
            bool bFound = false;
            DataTable dataTable = new DataTable();
            try
            {
                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == false)
                {
                    OdbcConnection db2Connection = new OdbcConnection();
                    //MessageBox.Show("GetDataTableByText_IsEDI: DB2ConnectionString: " + DB2ConnectionString);
                    db2Connection.ConnectionString = DB2ConnectionString; 

                    if (db2Connection.State != ConnectionState.Open)
                    {
                        db2Connection.Open();

                        iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", db2Connection);

                    }

                }

                OdbcConnection connection = (OdbcConnection)iQPUBLIC.PublicComponents.htMyVariable["DB2Conn"];
                if (connection.State != ConnectionState.Open)
                {
                   // MessageBox.Show("GetDataTableByText_IsEDI: DB2ConnectionString1: " + DB2ConnectionString);
                    connection.ConnectionString = DB2ConnectionString;  
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

                            if (dataTable.Rows.Count > 0)
                            {
                                bFound = true;
                            }

                        }
                    }
                }
                
            }
            catch (Exception ex)
            {
                //MessageBox.Show("GetDataTableByText_IsEDI: "+ ex.Message);

                throw;
               // WriteLog("IsEDI_Document: " + ex.Message, LogType.E);
                //MessageBox.Show(ex.Message);
                // return null;
                // dataTable = null;
               // bFound = false;
            }
            finally
            {
                if (dataTable != null)
                {
                    dataTable.Dispose();
                }
            }
            return bFound;
        }
        public static ServiceResponse GetDataTableFromSP(string connectionString, string storedProcedureName, SqlParameter[] parameters = null)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            DataTable dataTable = new DataTable();
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))

                using (SqlCommand cmd = new SqlCommand(storedProcedureName, connection))
                {
                    if (connection != null && !string.IsNullOrEmpty(storedProcedureName))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters);
                        }
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dataTable);
                        }
                    }
                }


                if (dataTable.Rows[0]["SuccessOrFail"].ToString() == "F")
                {
                    serviceResponse.Status = ServiceResponseStatus.Failed;
                }
                else
                {
                    serviceResponse.Status = ServiceResponseStatus.Success;
                    serviceResponse.Response = dataTable;
                }

            }
            catch (Exception ex)
            {
                serviceResponse.Status = ServiceResponseStatus.Failed;
            }
            finally
            {
                if (dataTable != null)
                {
                    dataTable.Dispose();
                    dataTable = null;
                }
            }

            return serviceResponse;
        }

        public static DataSet GetDataTableFromSP_As400(string connectionString, string storedProcedureName, SqlParameter[] parameters = null)
        {
            
            DataSet dataSet = new DataSet();

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(storedProcedureName, connection))
                    {
                        if (connection != null && !string.IsNullOrEmpty(storedProcedureName))
                        {
                            cmd.CommandType = CommandType.StoredProcedure;
                            if (parameters != null)
                            {
                                cmd.Parameters.AddRange(parameters);
                            }
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(dataSet);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                //serviceResponse.Status = ex.Message.ToString();
            }
            return dataSet;
        }

        public static ServiceResponse InsertOrUpdateBySP(string connectionString, string storedProcedureName, SqlParameter[] parameters)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            DataTable dataTable = new DataTable();
            try
            {
                //using (SqlConnection connection = new SqlConnection(connectionString))

                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("SQLConn") == false)
                {
                    SqlConnection sqlConnection = new SqlConnection();
                    sqlConnection.ConnectionString = connectionString;

                    if (sqlConnection.State != ConnectionState.Open)
                    {
                        sqlConnection.Open();
                        iQPUBLIC.PublicComponents.htMyVariable.Add("SQLConn", sqlConnection);
                    }
                }

                SqlConnection connection = (SqlConnection)iQPUBLIC.PublicComponents.htMyVariable["SQLConn"];
                if (connection.State != ConnectionState.Open)
                {
                    connection.ConnectionString = connectionString;
                    connection.Open();

                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("SQLConn") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("SQLConn");
                    }
                    iQPUBLIC.PublicComponents.htMyVariable.Add("SQLConn", connection);
                }

                using (SqlCommand cmd = new SqlCommand(storedProcedureName, connection))
                {
                    if (connection != null && !string.IsNullOrEmpty(storedProcedureName) && parameters != null)
                    {
                        
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddRange(parameters);
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dataTable);
                        }
                    }
                }


                if (dataTable.Rows[0]["SuccessOrFail"].ToString() == "F")
                {
                    serviceResponse.Status = ServiceResponseStatus.Failed;
                }
                else
                {
                    serviceResponse.Status = ServiceResponseStatus.Success;
                    serviceResponse.Response = dataTable;
                }
            }
            catch (Exception ex)
            {
                
                throw;
            }
            finally
            {
                if (dataTable != null)
                {
                    dataTable.Dispose();
                    dataTable = null;
                }
            }

            return serviceResponse;
        }

        public static ServiceResponse InsertOrUpdateBySP_BackUp(string connectionString, string storedProcedureName, SqlParameter[] parameters)
        {
            ServiceResponse serviceResponse = new ServiceResponse();
            DataTable dataTable = new DataTable();
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand cmd = new SqlCommand(storedProcedureName, connection))
                {
                    if (connection != null && !string.IsNullOrEmpty(storedProcedureName) && parameters != null)
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddRange(parameters);
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dataTable);
                        }
                    }
                }
                if (dataTable.Rows[0]["SuccessOrFail"].ToString() == "F")
                {
                    serviceResponse.Status = ServiceResponseStatus.Failed;
                }
                else
                {
                    serviceResponse.Status = ServiceResponseStatus.Success;
                    serviceResponse.Response = dataTable;
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            return serviceResponse;
        }


        public static void DisposeConnection_SQL(this SqlConnection connectionObject)
        {
            if (connectionObject != null)
            {
                connectionObject.Close();
                connectionObject.Dispose();
            }
        }

        public static void DisposeConnection_DB2(this OdbcConnection connectionObject)
        {
            if (connectionObject != null)
            {
                connectionObject.Close();
                connectionObject.Dispose();
            }
        }


        public static void DB2Conn_Open()
        {

            if (db2Connection.State != ConnectionState.Open)
            {
                db2Connection.Open();

                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == true)
                {
                    iQPUBLIC.PublicComponents.htMyVariable.Remove("DB2Conn");
                }

                iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", db2Connection);
            }
        }

        public static void SQLConn_Open()
        {

            if (sqlConnection.State != ConnectionState.Open)
            {
                sqlConnection.Open();
                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("SQLConn") == true)
                {
                    iQPUBLIC.PublicComponents.htMyVariable.Remove("SQLConn");
                }
                iQPUBLIC.PublicComponents.htMyVariable.Add("SQLConn", sqlConnection);

            }
        }


        public static void UpdateIntoFRPD010_UpdateIntoFRP001_InsertIntoFRP002(DataTable dtFRPD010, Dictionary<string, string> dicHeaderData_FRP001, DataTable dtFRP002TRU, int Linenumber, List<int> _listLinenumber_FRP002TRU, string connectionString, string ProNumber, OdbcParameter[] parameters = null)
        {

            try
            {
                string queryString = "select FHPRO, FHSTAT from FRP001 where FHPRO = " + ProNumber + " and FHSTAT = 'PP' ";

                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == false)
                {
                    OdbcConnection db2Connection = new OdbcConnection();
                    //MessageBox.Show("GetDataTableByText_IsEDI: DB2ConnectionString: " + DB2ConnectionString);
                    db2Connection.ConnectionString = DB2ConnectionString;

                    if (db2Connection.State != ConnectionState.Open)
                    {
                        db2Connection.Open();

                        iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", db2Connection);

                    }

                }

                OdbcConnection connection = (OdbcConnection)iQPUBLIC.PublicComponents.htMyVariable["DB2Conn"];
                if (connection.State != ConnectionState.Open)
                {
                    // MessageBox.Show("GetDataTableByText_IsEDI: DB2ConnectionString1: " + DB2ConnectionString);
                    connection.ConnectionString = DB2ConnectionString;
                    connection.Open();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("DB2Conn");
                    }
                    iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", connection);
                }

                bool ISFHSTAT_PP = false;

                #region"CHECK FHSTAT = 'PP'"
                DataTable dataTable = new DataTable();
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

                            if (dataTable.Rows.Count > 0)
                            {
                                ISFHSTAT_PP = true;
                            }

                        }
                    }
                }
                #endregion

                if (ISFHSTAT_PP == true)
                {

                }
                else
                {

                    //BOLDocumentDataExtraction  WriteLog(documentName + ":" + " Billed in AS400", LogType.E);
                    //InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " Billed in AS400", executionServiceName, "PerformDataExtraction_KM2");

                }

            }
            catch (Exception ex)
            {

            }

        }
    }
}
