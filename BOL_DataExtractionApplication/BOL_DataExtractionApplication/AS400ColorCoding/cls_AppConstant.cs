using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using BOLDocumentDataExtraction.CommonModule;

namespace BOLDocumentDataExtraction.AS400ColorCoding
{
   public static class cls_AppConstant
    {
        public static string Comma = ",";
        public static string Colons = ":";
        public static string SemiColons = ";";
        public static int RecordLength = 10;
        public static string[] StaticFieldNames ={ "PO","BL","SN"};
        
        public static string Inv_Value_decimalPoint = ".00";
        public static string DollerSign = "$";
        public static string Inv_Value_OverValue = "250,000";
        public static string Inv_Value_OverFinalValue = "1.00";
        public static string DotSign = ".";
        public static int DescriptionTextLength = 35;
        public static string Hash = "#";
        public static char HashSplit = '#';
        public static char Till = '~';
        public static string COD = "COD";
        public static string INV = "Inv_Value";
        public static string Del_Quote_No = "Del_Quote_No";
        
        public static string [] Punctuation = {";",":",",","." };
        public static string Replacewithspace = " ";
        public static string Del_Quote_Code = "QUOTE";
        public static string ERT_Code = "ERT_Code";
        public static string EMER = "EMER";
        public static string Tel_HazMat = "Tel_HazMat";
        public static string Con_TelNo = "Con_TelNo";
        public static string Con_TelNo_Code = "CCPHON";
        public static string FieldColorOne = "1";
        public static string FieldColorZero = "0";
        public static string FolderPath = "\\AS400FilePath";

        public static string PROFilePath = "\\AS400FilePath\\PRO.txt";
        public static string LogPath = "\\AS400FilePath\\LogFiles\\";
        public static string ProcessLogfile = "\\ProcessLog.txt"; 
        public static string ErrorLogFile = "\\ErrorLog.txt";
        public static string[] CODInvalidValue = { "$0.00","000","0.00","00","0.00"} ;//1230.00
        public static string[] ClassCodeArray = { ".0",".00" };
        public static string HazordousX = "X";
        public static string EMERDescription = "EMERGENCY RESPONSE NUMBER-";
        public static string DoublePipe = "||";
        public static string PO_BL_SN_NA = "NA";
        public static string PO_BL_SN_NS = "NS";
        public static string TruBot = "TRUBOT";

        public static string _ProNumber = "";
        public static string _ConnString = "";

        public static int AddressLength = 30;
        public static int PO_BL_SP_length = 20;
        public static int Shp_Con_Bill_Code_length = 7;
        public static int City_length = 20;
        public static int State_length = 2;
        public static int ZipCode_length = 6;
        public static int Terms_length = 3;
        public static int Total_Pieces_length = 5;
        public static int Total_Weight_length = 7;
        public static int RADF_RADT_length = 7;
        public static int COD_length = 9;
        public static int Dlv_length = 1;


        public static Dictionary<string, string> dicHdFldMapping
        { get; set; }
        public static Dictionary<string, string> dicLineFieldsMapping
        { get; set; }



        public static List<string> HeaderToLineItem()
        {
            try
            {
                List<string> list_temp = new List<string>();
                list_temp.Add("Special_Instr");
                list_temp.Add("Accessorial");
                list_temp.Add("Del_Requirements");
               
                return list_temp;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static List<KeyValuePair<String, string>> CODCashArray()
        {
            List<KeyValuePair<String, String>> _keyValuePairs = new List<KeyValuePair<String, String>>();
            try
            {

                _keyValuePairs.Add(new KeyValuePair<String, String>("FEE", "CODE FEE"));
                _keyValuePairs.Add(new KeyValuePair<String, String>("", ""));
                //_keyValuePairs.Add(new KeyValuePair<String, String>("PO_No", "XPO"));

            }
            catch (Exception)
            {

                throw;
            }
            return _keyValuePairs;
        }

        public static List<KeyValuePair<String, string>> MultiFieldValueFieldsName()
        {
            List<KeyValuePair<String, String>> _keyValuePairs = new List<KeyValuePair<String, String>>();
            try
            {

                _keyValuePairs.Add(new KeyValuePair<String, String>(cls_HeaderStaticField.Bol_No, "XBL"));
                _keyValuePairs.Add(new KeyValuePair<String, String>(cls_HeaderStaticField.Shp_No, "XSN"));
                _keyValuePairs.Add(new KeyValuePair<String, String>(cls_HeaderStaticField.PO_No, "XPO"));

            }
            catch (Exception)
            {

                throw;
            }
            return _keyValuePairs;
        }

        public static List<KeyValuePair<String, string>> MultiFieldValueRowColumnAS400()
        {
            List<KeyValuePair<String, String>> _keyValuePairs = new List<KeyValuePair<String, String>>();
            try
            {

                _keyValuePairs.Add(new KeyValuePair<String, String>(cls_HeaderStaticField.Bol_No, "13,5,20"));
                _keyValuePairs.Add(new KeyValuePair<String, String>(cls_HeaderStaticField.Shp_No, "14,41,20"));
                _keyValuePairs.Add(new KeyValuePair<String, String>(cls_HeaderStaticField.PO_No, "14,5,20"));


            }
            catch (Exception)
            {

                throw;
            }
            return _keyValuePairs;
        }

        public static List<string> AddressList()
        {
            List< String> _listAddress = new List< String>();
            try
            {
                _listAddress.Add(cls_HeaderStaticField.Shp_Addr_Line);
                _listAddress.Add(cls_HeaderStaticField.Con_Addr_Line);
                _listAddress.Add(cls_HeaderStaticField.Bill_To_Addr_Line);


            }
            catch (Exception)
            {

                throw;
            }
            return _listAddress;
        }

        /// <summary>
        /// Field list mapping with AS4000 color code fields
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> _HeaderFieldsMappingToAS400()
        {
            Dictionary<string, string> dicHeaderFieldsMapping = new Dictionary<string, string>();
            try
            {
                dicHeaderFieldsMapping.Add("BOL_No", "FHBL");
                dicHeaderFieldsMapping.Add("Shp_No", "FHSNUM");
                dicHeaderFieldsMapping.Add("PO_No", "FHPO");
                dicHeaderFieldsMapping.Add("Shp_Code", "FHSCD");
                dicHeaderFieldsMapping.Add("Shp_Name", "FHSNM");
                dicHeaderFieldsMapping.Add("Shp_Addr_Line", "FHSA1");
                dicHeaderFieldsMapping.Add("Shp_Addr_Line1", "FHSA2");
                dicHeaderFieldsMapping.Add("Shp_City", "FHSCT");
                dicHeaderFieldsMapping.Add("Shp_State", "FHSST");
                dicHeaderFieldsMapping.Add("Shp_ZipCode", "FHSZIP");
                dicHeaderFieldsMapping.Add("Con_Code", "FHCCD");
                dicHeaderFieldsMapping.Add("Con_Name", "FHCNM");
                dicHeaderFieldsMapping.Add("Con_Addr_Line", "FHCA1");
                dicHeaderFieldsMapping.Add("Con_Addr_Line1", "FHCA2");
                
                dicHeaderFieldsMapping.Add("Con_City", "FHCCT");
                dicHeaderFieldsMapping.Add("Con_State", "FHCST");
                dicHeaderFieldsMapping.Add("Con_ZipCode", "FHCZIP");
                dicHeaderFieldsMapping.Add("Bill_To_Code", "FHBTC");
                dicHeaderFieldsMapping.Add("Bill_To_Name", "FHBNM");
                dicHeaderFieldsMapping.Add("Bill_To_Addr_Line", "FHBA1");
                dicHeaderFieldsMapping.Add("Bill_To_Addr_Line1", "FHBA2");
                dicHeaderFieldsMapping.Add("Bill_To_City", "FHBCT");
                dicHeaderFieldsMapping.Add("Bill_To_State", "FHBST");
                dicHeaderFieldsMapping.Add("Bill_To_ZipCode", "FHBZIP");
                dicHeaderFieldsMapping.Add("Terms", "FHTRM");
                dicHeaderFieldsMapping.Add("Total_Pieces", "FHTOTP");
                dicHeaderFieldsMapping.Add("Total_Weight", "FHSWGT");
                dicHeaderFieldsMapping.Add("Dates_RADF", "FHRADF");
                dicHeaderFieldsMapping.Add("Dates_MABD", "FHRADT");
                dicHeaderFieldsMapping.Add("COD", "FHCOD");
                dicHeaderFieldsMapping.Add("Dlv", "FHDC");
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog(" cls_AS400ColorCode._HeaderFieldsMappingToAS400 -> " + ex.Message.ToString(), LogType.E);
            }
            return dicHeaderFieldsMapping;
        }  
        public static Dictionary<string, string> _LineFieldsMappingToAS400()
        {
            Dictionary<string, string> dicLineFieldsMapping = new Dictionary<string, string>();
            try
            {
                dicLineFieldsMapping.Add("Quantity", "FDPCS");
                dicLineFieldsMapping.Add("Class_Code", "FDCMCL");
                dicLineFieldsMapping.Add("HazMat_Code", "FDHAZ");
                dicLineFieldsMapping.Add("Pallate_Code", "FDPKGC");
                dicLineFieldsMapping.Add("Description", "FDDES");
                dicLineFieldsMapping.Add("NMFC", "FDDES"); //TBD
                dicLineFieldsMapping.Add("Weight", "FDWGT");
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog(" cls_AS400ColorCode._LineFieldsMappingToAS400 -> " + ex.Message.ToString(),LogType.E);
            }
            return dicLineFieldsMapping;
        }

        public static Dictionary<string, string> _LineFieldsMappingForTableInsert()
        {
            Dictionary<string, string> dicLineFieldsMapping = new Dictionary<string, string>();
            try
            {
              
                dicLineFieldsMapping.Add("Quantity", "FDPCS");
                dicLineFieldsMapping.Add("Class_Code", "FDCMCL");
                dicLineFieldsMapping.Add("HazMat_Code", "FDHAZ");
                dicLineFieldsMapping.Add("Pallate_Code", "FDPKGC");
                dicLineFieldsMapping.Add("Description", "FDDES");
                dicLineFieldsMapping.Add("NMFC", "FDDES"); //TBD
                dicLineFieldsMapping.Add("Weight", "FDWGT");
                dicLineFieldsMapping.Add("LineNumber", "FDLI");
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog(" cls_AS400ColorCode._LineFieldsMappingToAS400 -> " + ex.Message.ToString(),LogType.E );
            }
            return dicLineFieldsMapping;
        }


        public static  DataTable FieldsMappingForTableFRPD010()
        {
            DataTable dtSortedTable = new DataTable();
            try
            {
                
                dtSortedTable.Columns.Add("ProNumber");
                dtSortedTable.Columns.Add("FileName");
                dtSortedTable.Columns.Add("FieldName");
                dtSortedTable.Columns.Add("LineNumber");
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog(" cls_AS400ColorCode.FieldsMappingForTableFRPD010 -> " + ex.Message.ToString(),LogType.E );
            }
            return dtSortedTable;
        }




    }
}
