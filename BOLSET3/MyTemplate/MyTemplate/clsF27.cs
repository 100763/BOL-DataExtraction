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

namespace BOLSET3IQ
{
    public class clsF27
    {
        ////string DBConnectionString = "Data Source=172.1.254.121,1433;Initial Catalog=SAIA_Extraction;User ID=sa;Password=Password1234;Connect Timeout=30";
        //string SQLConnectionString = "Data Source=DCAP8SQLDEV01;Initial Catalog=SAIABOLTransaction;User ID=boladmin;Password=bolSIocr@65;Connect Timeout=30"; //   SAIA@DBTest01
        //string DB2ConnectionString = "Driver={iSeries Access ODBC Driver};System=SAIATEST;UID=BOTBOLOCR;PWD=saia789$;Naming=1;DBQ=*USRLIBL;Compression=1;";
        string DB2ConnectionString = string.Empty;
        string SQLConnectionString = string.Empty;
        public RetStructF27 F27(string strRoutineNo, string strArg, string strFldId, string currRowNo, string type)
        {
            //MessageBox.Show("F27 function reached");
           //RetStructF27 objStructF27 = new RetStructF27();
            Module1.objStructF27.Status = "F";
           
            string TotalWeight_TotalPieces = "";
            string ShipperName_ShipperCode = "";
            try
            {            
                switch (strRoutineNo)
                {
                    case "CKM250_F27":  //Total Weight _DB from server
                        try
                        {
                            string PRONo = Path.GetFileNameWithoutExtension(iQPUBLIC.PublicComponents.PrimaryImagePath);
                            PRONo = PRONo.Split('_')[PRONo.Split('_').Length - 1];

                            string total_Weight = "";
                            string total_Pieses = "";

                            if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET3_TotalWeight"))
                            {
                                iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET3_TotalWeight");
                            }
                            if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET3_TotalPieces"))
                            {
                                iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET3_TotalPieces");
                            }

                            DataTable dtPickup_Consigne_Details = GetPickupConsigneDetails_forPCS_Weight(PRONo);
                            //Module1.dtPickup_Consigne_Details = GetPickupConsigneDetails_forPCS_Weight(PRONo);
                            if (dtPickup_Consigne_Details != null && dtPickup_Consigne_Details.Rows.Count > 0)
                            {
                                if (!DBNull.Value.Equals(dtPickup_Consigne_Details.Rows[0]["FHSWGT"]))
                                {
                                    total_Weight = dtPickup_Consigne_Details.Rows[0]["FHSWGT"].ToString().Trim();
                                }
                                iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET3_TotalWeight", total_Weight);

                                if (!DBNull.Value.Equals(dtPickup_Consigne_Details.Rows[0]["FHTOTP"]))
                                {
                                    total_Pieses =dtPickup_Consigne_Details.Rows[0]["FHTOTP"].ToString().Trim();
                                }
                                iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET3_TotalPieces", total_Pieses);
                            }


                            TotalWeight_TotalPieces = total_Weight + "#" + total_Pieses;

                            Module1.objStructF27.Status = "S";
                            Module1.objStructF27.ContentsOfField = TotalWeight_TotalPieces;
                            Module1.objStructF27.ConfLevel = 90;
                            Module1.objStructF27.ConfString = "9".PadLeft(Module1.objStructF27.ContentsOfField.Length, '9');
                            Module1.objStructF27.X1 = 5;
                            Module1.objStructF27.X2 = 5;
                            Module1.objStructF27.Y1 = 10;
                            Module1.objStructF27.Y2 = 10;

                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("F27 Error in Special Instruction table Fetching - " + ex.Message.ToString());
                        }

                        break;
                    case "CKM250":  //Total Weight _DB
                        try
                        {
                            //commented for Rajesh Sir Server Testing 

                            string total_Weight = "";
                            string total_Pieses = "";
                            if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET3_TotalWeight"))
                            {
                                total_Weight = (string)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_TotalWeight"];
                            }

                            if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET3_TotalPieces"))
                            {
                                total_Pieses = (string)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_TotalPieces"];
                            }
                        
                            TotalWeight_TotalPieces = total_Weight + "#" + total_Pieses;

                            Module1.objStructF27.Status = "S";
                            Module1.objStructF27.ContentsOfField = TotalWeight_TotalPieces;
                            Module1.objStructF27.ConfLevel = 90;
                            Module1.objStructF27.ConfString = "9".PadLeft(Module1.objStructF27.ContentsOfField.Length, '9');
                            Module1.objStructF27.X1 = 5;
                            Module1.objStructF27.X2 = 5;
                            Module1.objStructF27.Y1 = 10;
                            Module1.objStructF27.Y2 = 10;
                           
                            // Special Instruction database data fetching
                            if (Module1.dtAccessorials.Rows.Count == 0)
                            {
                                // Module1.dtAccessorials = GetSpecialInstructDetails_AccessorialDetails("USP_ACCESSORIALS_Getdata");
                                Module1.dtAccessorials = GetSpecialInstructDetails_AccessorialDetails("USP_SpecialInstruction_Getdata");
                            }

                            if (Module1.Dt_IgnoreShipperDetails.Rows.Count == 0)
                            {
                                Module1.Dt_IgnoreShipperDetails = GetIgnoreShippersDetails();
                            }


                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("F27 Error in Special Instruction table Fetching - " + ex.Message.ToString());
                        }

                        break;
                        #region shipper Name
                        ////case "CKM_SI":  //Shipper Name and COde
                        ////    try
                        ////    {

                        ////        string Shipper_Name = "";
                        ////        string Shipper_Code = "";
                        ////        if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET3_ShipperName"))
                        ////        {
                        ////            Shipper_Name = (string)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_ShipperName"];
                        ////        }

                        ////        if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET3_ShipperCode"))
                        ////        {
                        ////            Shipper_Code = (string)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_ShipperCode"];
                        ////        }

                        ////        ShipperName_ShipperCode = Shipper_Name + "#" + Shipper_Code;

                        ////        Module1.objStructF27.Status = "S";
                        ////        Module1.objStructF27.ContentsOfField = ShipperName_ShipperCode;
                        ////        Module1.objStructF27.ConfLevel = 90;
                        ////        Module1.objStructF27.ConfString = "9".PadLeft(Module1.objStructF27.ContentsOfField.Length, '9');
                        ////        Module1.objStructF27.X1 = 5;
                        ////        Module1.objStructF27.X2 = 5;
                        ////        Module1.objStructF27.Y1 = 10;
                        ////        Module1.objStructF27.Y2 = 10;

                        ////        if (Module1.Dt_IgnoreShipperDetails.Rows.Count == 0)
                        ////        {
                        ////            Module1.Dt_IgnoreShipperDetails = GetIgnoreShippersDetails();
                        ////        }



                        ////    }
                        ////    catch (Exception ex)
                        ////    {
                        ////        //MessageBox.Show("F27 Error in Special Instruction table Fetching - " + ex.Message.ToString());
                        ////    }

                        ////break;
                        #endregion

                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("F27 - " + ex.Message.ToString());
            }
            return Module1.objStructF27;
            //return objStructF27;
        }

        private DataTable GetPickupConsigneDetails_forPCS_Weight(string PRONumber)
        {
            DataTable dtPickupConsigneDetails = new DataTable();
            try
            {
                ////if (string.IsNullOrEmpty(iQPUBLIC.PublicComponents.DbConnectionString1))
                ////{
                ////    iQPUBLIC.PublicComponents.DbConnectionString1 = ConfigurationManager.ConnectionStrings["DB2Conn"].ToString();  //DB2
                ////}

                if (!string.IsNullOrEmpty(PRONumber))
                {
                    string strquery = "select FHSWGT, FHTOTP from FRP001 where FHPRO= '" + PRONumber + "'";

                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == false)
                    {
                        OdbcConnection db2Connection = new OdbcConnection();
                        //db2Connection.ConnectionString = iQPUBLIC.PublicComponents.DbConnectionString1;//DB2ConnectionString;
                        if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_DB2Conn") == true)
                        {
                            DB2ConnectionString=(string) iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_DB2Conn"];
                        }
                        else
                        {
                            DB2ConnectionString = ConfigurationManager.ConnectionStrings["AS400DBConnString"].ToString();

                        }


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
                        if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_DB2Conn") == true)
                        {
                            DB2ConnectionString = (string)iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_DB2Conn"];
                        }
                        else
                        {
                            DB2ConnectionString = ConfigurationManager.ConnectionStrings["AS400DBConnString"].ToString();

                        }
                        ////connection.ConnectionString = iQPUBLIC.PublicComponents.DbConnectionString1;
                        connection.ConnectionString = DB2ConnectionString;
                        connection.Open();
                        if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == true)
                        {
                            iQPUBLIC.PublicComponents.htMyVariable.Remove("DB2Conn");
                        }
                        iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", connection);
                    }


                    using (OdbcCommand cmd = new OdbcCommand(strquery, connection))
                    {
                        if (connection != null && !string.IsNullOrEmpty(strquery))
                        {
                            using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                            {
                                da.Fill(dtPickupConsigneDetails);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("GetPickupConsigneDetails_forPCS_Weight: " + ex.Message);

            }

            return dtPickupConsigneDetails;
        }

        private DataTable GetSpecialInstructDetails_AccessorialDetails(string spName)
        {
            DataTable dtSpecial_Instruct_Accessorial = new DataTable();
            try
            {
               

                ////if (string.IsNullOrEmpty(iQPUBLIC.PublicComponents.DbConnectionString2))
                ////{
                ////    iQPUBLIC.PublicComponents.DbConnectionString = ConfigurationManager.ConnectionStrings["SQLConn"].ToString(); //SQL
                ////}

                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("SQLConn") == false)
                {
                    SqlConnection sqlConnection = new SqlConnection();
                    ////sqlConnection.ConnectionString = iQPUBLIC.PublicComponents.DbConnectionString;//SQLConnectionString;
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_SQLConn") == true)
                    {
                        SQLConnectionString = (string)iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_SQLConn"];
                    }
                    else
                    {
                        SQLConnectionString = ConfigurationManager.ConnectionStrings["DBPath"].ToString();

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
                    ////connection.ConnectionString = iQPUBLIC.PublicComponents.DbConnectionString;
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_SQLConn") == true)
                    {
                        SQLConnectionString = (string)iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_SQLConn"];
                    }
                    else
                    {
                        SQLConnectionString = ConfigurationManager.ConnectionStrings["DBPath"].ToString();

                    }
                    connection.ConnectionString = SQLConnectionString;
                    connection.Open();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("SQLConn") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("SQLConn");
                    }
                    iQPUBLIC.PublicComponents.htMyVariable.Add("SQLConn", connection);
                }

                //string spName = "USP_Special_Instruct_GetData";

                //using (SqlConnection connection = new SqlConnection(connectionString))

                using (SqlCommand cmd = new SqlCommand(spName, connection))
                {
                    if (connection != null && !string.IsNullOrEmpty(spName))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dtSpecial_Instruct_Accessorial);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("GetSpecialInstructDetails: " + ex.Message);
            }
            return dtSpecial_Instruct_Accessorial;
        }

        private DataTable GetIgnoreShippersDetails()
        {
            DataTable dtIgnoreSIShipper = new DataTable();
            try
            {


                ////if (string.IsNullOrEmpty(iQPUBLIC.PublicComponents.DbConnectionString2))
                ////{
                ////    iQPUBLIC.PublicComponents.DbConnectionString = ConfigurationManager.ConnectionStrings["SQLConn"].ToString(); //SQL
                ////}

                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("SQLConn") == false)
                {
                    SqlConnection sqlConnection = new SqlConnection();
                    ////sqlConnection.ConnectionString = iQPUBLIC.PublicComponents.DbConnectionString;//SQLConnectionString;
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_SQLConn") == true)
                    {
                        SQLConnectionString = (string)iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_SQLConn"];
                    }
                    else
                    {
                        SQLConnectionString = ConfigurationManager.ConnectionStrings["DBPath"].ToString();

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
                    ////connection.ConnectionString = iQPUBLIC.PublicComponents.DbConnectionString;
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_SQLConn") == true)
                    {
                        SQLConnectionString = (string)iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_SQLConn"];
                    }
                    else
                    {
                        SQLConnectionString = ConfigurationManager.ConnectionStrings["DBPath"].ToString();

                    }
                    connection.ConnectionString = SQLConnectionString;
                    connection.Open();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("SQLConn") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("SQLConn");
                    }
                    iQPUBLIC.PublicComponents.htMyVariable.Add("SQLConn", connection);
                }

                string query = "select * from IgnoreSpecislInstructionCustomer";

                //using (SqlConnection connection = new SqlConnection(connectionString))

                using (SqlCommand cmd = new SqlCommand(query, connection))
                {
                    if (connection != null && !string.IsNullOrEmpty(query))
                    {
                        //cmd.CommandType = CommandType.StoredProcedure;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dtIgnoreSIShipper);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("GetSpecialInstructDetails: " + ex.Message);
            }
            return dtIgnoreSIShipper;
        }

    }
}