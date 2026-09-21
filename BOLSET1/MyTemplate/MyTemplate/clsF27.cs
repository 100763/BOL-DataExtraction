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
using BarcodeReader;
using System.IO;

namespace BOLSET1IQ
{
    public class clsF27
    {

        public RetStructF27 F27(string strRoutineNo, string strArg, string strFldId, string currRowNo, string type)
        {
            //RetStructF27 objStructF27 = new RetStructF27();
            //objStructF27.Status = "F";
            Module1.objStructF27.Status = "F";

            //if (Module1.imageName.ToLower() != Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath).ToLower())
            //{
            //    //Module1.bFoundSearchKeywordTable = false;
            //    Module1.func_ReadBarcode();
            //    Module1.imageName = Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath);
            //}

            try
            {
                switch (strRoutineNo)
                {
                    case "CKM222":
                        Module1.func_ReadBarcode();
                        if (Module1.Barcodes != null)
                        {
                            if (Module1.Barcodes.Count > 0)
                            {
                                Module1.objStructF27.ContentsOfField = Module1.Barcodes[0];
                            }
                            else
                            {
                                //Module1.objStructF27.ContentsOfField = "*****";
                                Module1.objStructF27.ContentsOfField = "*****";
                            }
                        }
                        else
                        {
                            //Module1.objStructF27.ContentsOfField = "*****";
                            Module1.objStructF27.ContentsOfField = "*****";
                        }
                        Module1.objStructF27.Status = "S";
                        Module1.objStructF27.X1 = 5;
                        Module1.objStructF27.Y1 = 5;
                        Module1.objStructF27.X2 = 10;
                        Module1.objStructF27.Y2 = 10;
                        Module1.objStructF27.PageNo = 999;
                        Module1.objStructF27.ConfLevel = 90;
                        //objStructF27.ConfString = "9".PadLeft(strPaymentTerm.Length, '9');

                        break;

                    case "CKM223":
                        //Module1.func_ReadBarcode();
                        if (Module1.Barcodes != null)
                        {
                            if (Module1.Barcodes.Count > 1)
                            {
                                Module1.objStructF27.ContentsOfField = Module1.Barcodes[1];
                            }
                            else
                            {
                                Module1.objStructF27.ContentsOfField = "*****";
                            }
                        }
                        else
                        {
                            Module1.objStructF27.ContentsOfField = "*****";
                        }
                        Module1.objStructF27.Status = "S";
                        Module1.objStructF27.X1 = 5;
                        Module1.objStructF27.Y1 = 5;
                        Module1.objStructF27.X2 = 10;
                        Module1.objStructF27.Y2 = 10;
                        Module1.objStructF27.PageNo = 999;
                        Module1.objStructF27.ConfLevel = 90;
                        break;

                    case "CKM224":
                        //Module1.func_ReadBarcode();
                        if (Module1.Barcodes != null)
                        {
                            if (Module1.Barcodes.Count > 2)
                            {

                                //objStructF27.ConfString = "9".PadLeft(strPaymentTerm.Length, '9');

                                Module1.objStructF27.ContentsOfField = Module1.Barcodes[2];
                            }
                            else
                            {
                                Module1.objStructF27.ContentsOfField = "*****";
                            }
                        }
                        else
                        {
                            Module1.objStructF27.ContentsOfField = "*****";
                        }
                        Module1.objStructF27.Status = "S";
                        Module1.objStructF27.X1 = 5;
                        Module1.objStructF27.Y1 = 5;
                        Module1.objStructF27.X2 = 10;
                        Module1.objStructF27.Y2 = 10;
                        Module1.objStructF27.PageNo = 999;
                        Module1.objStructF27.ConfLevel = 90;
                        break;
                }
            }
            catch (Exception )
            {
                //MessageBox.Show("F27 - " + ex.Message.ToString());
            }
            return Module1.objStructF27;
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
            catch (Exception)
            {
                //MessageBox.Show("GetValueFromDB: " + ex.Message);

            }

            return PaymentValue;

        }


       
    }
}