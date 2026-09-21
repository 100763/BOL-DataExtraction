using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using iQDataProvider;
using iQPUBLIC;
//using System.Windows.Forms;
using System.Data.SqlClient;

namespace BOLSET4IQ
{
    public class clsF27
    {
        public RetStructF27 F27(string strRoutineNo, string strArg, string strFldId, string currRowNo, string type)
        {
            RetStructF27 objStructF27 = new RetStructF27();
            objStructF27.Status = "F";
                       
            try
            {
                switch (strRoutineNo)
                {
                    case "C75":


                        string strPaymentTerm = GetValueFromDB(strArg);

                        if (strPaymentTerm.Trim() != "")
                        {
                            objStructF27.Status = "S";
                            objStructF27.X1 = 5;
                            objStructF27.Y1 = 5;
                            objStructF27.X2 = 10;
                            objStructF27.Y2 = 10;
                            objStructF27.PageNo = 999;
                            objStructF27.ConfLevel = 90;
                            objStructF27.ConfString = "9".PadLeft(strPaymentTerm.Length, '9');

                            objStructF27.ContentsOfField = strPaymentTerm;


                        }
                        else
                        {
                            objStructF27.Status = "F";
                        }


                        break;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("F27 - " + ex.Message.ToString());
            }
            return objStructF27;
        }

        private string GetValueFromDB(string strArg)
        {
            string PaymentValue = "";
            try
            {
                if (strArg.Trim() != "")
                {

                    DataTable dt = iQPUBLIC.Common.GetDataTable();


                    string InvValue = iQPUBLIC.Common.GetHeaderDataForCell(1, dt);
                    string connstr = strArg;
                    string strquery = "select Payment_Term from VendorDetails where Invoice_Number= '" + InvValue + "'";

                    SqlConnection sqlclonn = new SqlConnection(connstr);
                    sqlclonn.Open();
                    SqlCommand sqlcmd = new SqlCommand(strquery, sqlclonn);

                    SqlDataReader oreader = sqlcmd.ExecuteReader();
                    if (oreader.HasRows)
                    {
                        while(oreader.Read())
                        {
                            PaymentValue = oreader.GetString(0);
                        }
                    }
                    else
                    {
                        PaymentValue = "";
                    }


                }
                else
                {
                    PaymentValue = "";
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("GetValueFromDB: " + ex.Message);

            }

            return PaymentValue;

        }


       
    }
}