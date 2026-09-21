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

namespace BOLSET3IQ
{
    public class clsF27_old
    {
        //string DBConnectionString = "Data Source=172.1.254.121,1433;Initial Catalog=SAIA_Extraction;User ID=sa;Password=Password1234;Connect Timeout=30";
        //string DBConnectionString = "Data Source = DCAP8SQLDEV01; Initial Catalog = SAIABOLTransaction; User ID = boladmin; Password=SAIA @DBTest01; Connect Timeout = 30";
        string SQLConnectionString = "Data Source=DCAP8SQLDEV01;Initial Catalog=SAIABOLTransaction;User ID=boladmin;Password=SAIA@DBTest01;Connect Timeout=30";
        string DB2ConnectionString = "Driver={iSeries Access ODBC Driver};System=SAIATEST;UID=BOTBOLOCR;PWD=saia789$;Naming=1;DBQ=*USRLIBL;Compression=1;";

        public RetStructF27 F27(string strRoutineNo, string strArg, string strFldId, string currRowNo, string type)
        {
            //MessageBox.Show("F27 function reached");
            RetStructF27 objStructF27 = new RetStructF27();
            objStructF27.Status = "F";
            string TotalWeight = "";
            string TotalWeight_TotalPieces = "";
            try
            {
                switch (strRoutineNo)
                {
                    case "CKM250":  //Total Weight
                        try
                        {
                            //commented for Rajesh Sir Server Testing 
                            string PRONo = Path.GetFileNameWithoutExtension(iQPUBLIC.PublicComponents.PrimaryImagePath);
                            string[] splitPro = PRONo.Split('_');
                            PRONo = splitPro[0];
                            iQPUBLIC.PublicComponents.htMyVariable = new Hashtable();
                            DataTable dtPickup_Consigne_Details = GetPickupConsigneDetails_forPCS_Weight(51, PRONo);
                            if (dtPickup_Consigne_Details != null && dtPickup_Consigne_Details.Rows.Count > 0)
                            {
                                iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET3_PickupConsigneDetails", dtPickup_Consigne_Details);

                                if (dtPickup_Consigne_Details != null && dtPickup_Consigne_Details.Rows.Count > 0)
                                {
                                    TotalWeight = dtPickup_Consigne_Details.Rows[0].Field<string>("PCWGT");
                                }
                                
                                objStructF27.Status = "S";
                                objStructF27.ContentsOfField = TotalWeight;
                                objStructF27.ConfLevel = 90;
                                objStructF27.ConfString = "9".PadLeft(objStructF27.ContentsOfField.Length, '9');
                                objStructF27.X1 = 5;
                                objStructF27.X2 = 5;
                                objStructF27.Y1 = 10;
                                objStructF27.Y2 = 10;
                            }


                            ///  // Special Instruction database data fetching

                            if (Module1.dtSpecial_Instruct.Rows.Count==0)
                            {
                                Module1.dtSpecial_Instruct = GetSpecialInstructDetails(DBConnectionString);
                                
                            }
                            // Module1.dtSpecial_Instruct = GetSpecialInstructDetails(DBConnectionString);
                            //if (Module1.dtSpecial_Instruct.Rows.Count>=0)
                            //{
                            //    //MessageBox.Show("Special Instruction total rows:"+ Module1.dtSpecial_Instruct.Rows.Count.ToString());
                            //}
                            if (Module1.dtAccessorials.Rows.Count == 0)
                            {
                                //Module1.dtSpecial_Instruct = GetSpecialInstructDetails(DBConnectionString);
                                Module1.dtAccessorials = GetaccessorialDetails(DBConnectionString);
                            }

                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("F27 Error in Special Instruction table Fetching - " + ex.Message.ToString());
                        }

                        break;

                    //case "CKM251": // Special Instructions
                    //    DataTable dtSpectial_Instruct = new DataTable();
                    //    dtSpectial_Instruct = GetSpecialInstructDetails(DBConnectionString);
                    //    break;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("F27 - " + ex.Message.ToString());
            }
            return objStructF27;
        }

        private DataTable GetPickupConsigneDetails_forPCS_Weight(long documentID, string PRONumber)
        {
            DataTable dtPickupConsigneDetails = new DataTable();
            try
            {
                if (!string.IsNullOrEmpty(PRONumber))
                {

                    string strquery = "select * from DSP092 where PCPRO= '" + PRONumber + "'";

                    // DataTable dataTable = new DataTable();
                    try
                    {
                        using (SqlConnection connection = new SqlConnection(DBConnectionString))
                        using (SqlCommand cmd = new SqlCommand(strquery, connection))
                        {
                            if (connection != null)
                            {
                                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                                {
                                    da.Fill(dtPickupConsigneDetails);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // MessageBox.Show("GetPickupConsigneDetails_forPCS_Weight: " + ex.Message);
                    }


                    //return dtPickupConsigneDetails;
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("GetPickupConsigneDetails_forPCS_Weight: " + ex.Message);
            }
            return dtPickupConsigneDetails;
        }

        private DataTable GetSpecialInstructDetails(string connectionString)
        {
            DataTable dtSpecial_Instruct = new DataTable();
            try
            {
                string spName = "USP_Special_Instruct_GetData";
                using (SqlConnection connection = new SqlConnection(connectionString))
                using (SqlCommand cmd = new SqlCommand(spName, connection))
                {
                    if (connection != null && !string.IsNullOrEmpty(spName))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dtSpecial_Instruct);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("GetSpecialInstructDetails: " + ex.Message);
            }
            return dtSpecial_Instruct;
        }

        private DataTable GetaccessorialDetails(string connectionString)
        {
            DataTable dtAccessorialsDetails = new DataTable();
            try
            {
                string strquery = "select * from ACCESSORIALS ";

                using (SqlConnection connection = new SqlConnection(DBConnectionString))
                using (SqlCommand cmd = new SqlCommand(strquery, connection))
                {
                    if (connection != null)
                    {
                        using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                        {
                            da.Fill(dtAccessorialsDetails);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("GetaccessorialDetails: " + ex.Message);
            }
            return dtAccessorialsDetails;
        }

        #region Extra code
        ////private string GetValueFromDB(string strArg)
        ////{
        ////    string PaymentValue = "";
        ////    try
        ////    {
        ////        if (strArg.Trim() != "")
        ////        {

        ////            DataTable dt = iQPUBLIC.Common.GetDataTable();


        ////            string InvValue = iQPUBLIC.Common.GetHeaderDataForCell(1, dt);
        ////            string connstr = strArg;
        ////            string strquery = "select Payment_Term from VendorDetails where Invoice_Number= '" + InvValue + "'";

        ////            SqlConnection sqlclonn = new SqlConnection(connstr);
        ////            sqlclonn.Open();
        ////            SqlCommand sqlcmd = new SqlCommand(strquery, sqlclonn);

        ////            SqlDataReader oreader = sqlcmd.ExecuteReader();
        ////            if (oreader.HasRows)
        ////            {
        ////                while(oreader.Read())
        ////                {
        ////                    PaymentValue = oreader.GetString(0);
        ////                }
        ////            }
        ////            else
        ////            {
        ////                PaymentValue = "";
        ////            }


        ////        }
        ////        else
        ////        {
        ////            PaymentValue = "";
        ////        }
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////       // MessageBox.Show("GetValueFromDB: " + ex.Message);

        ////    }

        ////    return PaymentValue;

        ////}

        //private DataTable GetPickupConsigneDetails(long documentID, string PRONumber)
        //{
        //    DataTable dtPickupConsigneDetails = new DataTable();
        //    try
        //    {
        //        if (!string.IsNullOrEmpty(PRONumber))
        //        {
        //            string spName = "USP_PickupDetail_DSP092_Getdata";
        //            SqlParameter headerFieldDataParameter = new SqlParameter("@PRONumber", PRONumber);
        //            SqlParameter[] parameters = new SqlParameter[] { headerFieldDataParameter };
        //            dtPickupConsigneDetails = GetDataTableFromSP(DBConnectionString, spName, parameters);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //       // MessageBox.Show("GetPickupConsigneDetails: " + ex.Message);
        //    }
        //    return dtPickupConsigneDetails;
        //}
        //private DataTable GetDataTableFromSP(string connectionString, string storedProcedureName, SqlParameter[] parameters = null)
        //{
        //    DataTable dataTable = new DataTable();
        //    try
        //    {
        //        using (SqlConnection connection = new SqlConnection(connectionString))
        //        using (SqlCommand cmd = new SqlCommand(storedProcedureName, connection))
        //        {
        //            if (connection != null && !string.IsNullOrEmpty(storedProcedureName))
        //            {
        //                cmd.CommandType = CommandType.StoredProcedure;
        //                if (parameters != null)
        //                {
        //                    cmd.Parameters.AddRange(parameters);
        //                }
        //                using (SqlDataAdapter da = new SqlDataAdapter(cmd))
        //                {
        //                    da.Fill(dataTable);
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //       // MessageBox.Show("GetDataTableFromSP: " + ex.Message);
        //    }
        //    return dataTable;
        //}
        #endregion





    }
}