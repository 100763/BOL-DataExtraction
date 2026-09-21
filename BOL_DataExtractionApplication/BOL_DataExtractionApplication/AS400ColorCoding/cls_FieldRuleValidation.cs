using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Reflection;
using System.ComponentModel;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using System.Data.Odbc;
using BOLDocumentDataExtraction.CommonModule;
using BOLDocumentDataExtraction.DAL;

namespace BOLDocumentDataExtraction.AS400ColorCoding
{
    public class cls_FieldRuleValidation
    {

        public static string[] Arry_MultipleFieldValue = null;
        public static string[] Array_MultiFieldColor = null;

        public static DataTable dtSortedTable = null;



        public static DataSet ApplyRulesOnFileds(DataTable dt_Header, DataTable dt_LineItem)
        {
            DataSet ds = new DataSet();
            try
            {
                try
                {
                    // ds = UpdateFieldColorInDatatable(dt_Header, dt_LineItem); // Optimization
                    // ds = RemoveStarFieldValue(ds.Tables["Table1"], ds.Tables["Table2"]);// Optimization 
                    ds = NMFC(dt_Header, dt_LineItem);
                    // ds = NMFC(ds.Tables["Table1"], ds.Tables["Table2"]);
                    ds = AddHeaderFieldInLineItem(ds.Tables["Table1"], ds.Tables["Table2"]);

                }
                catch (Exception ex)
                {
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in ApplyRulesOnFileds submethod (AddHeaderFieldInLineItem) " + ex,LogType.E );
                }
                try
                {
                    ds = SortMultiHeaderFieldValue(ds.Tables["Table1"], ds.Tables["Table2"]);
                }
                catch (Exception ex)
                {
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in ApplyRulesOnFileds subMethod(SortMultiHeaderFieldValue)" + ex,LogType.E );
                }
                try
                {

                    ds = CODCashPaymentOrNot(ds);
                    ds = TelephoneHazmatInLineItem(ds.Tables["Table1"], ds.Tables["Table2"]);
                    ds = TelephoneConsiInLineItem(ds.Tables["Table1"], ds.Tables["Table2"]);
                    ds = FindDescription(ds);


                }
                catch (Exception ex)
                {
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in ApplyRulesOnFileds subMethod(CODCashPaymentOrNot)" + ex,LogType.E );
                }
                return ds;
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exeption in ApplyRulesOnFileds--->" + ex,LogType.E );
                return ds;
            }
        }





        #region For BOL,PO,SN

        /// <summary>
        ///if BOL is empty then replace load no value in BOL field
        ///if BOL is not empty then Load number field value append in line item
        /// </summary>
        /// <returns></returns>
        public static DataSet CheckLoadNoForBOL(DataTable dt_header, DataTable dt_lineitem)
        {
            DataTable dt_bol = new DataTable();
            DataTable dt_LoadNo = new DataTable();
            DataSet ds = new DataSet();
            string _BOl_No = string.Empty;
            try
            {
                dt_bol = FindFieldRowInDataTable(dt_header, cls_HeaderStaticField.Bol_No);
                dt_LoadNo = FindFieldRowInDataTable(dt_header, cls_HeaderStaticField.load_no);

                if (dt_LoadNo.Rows.Count > 0)
                {
                    _BOl_No = Convert.ToString(dt_bol.Rows[0]["FieldValue"].ToString());
                    if (!string.IsNullOrEmpty(_BOl_No))
                    {
                        //if bol number is not null then add load number in line item 
                        dt_lineitem = AddFieldValueinLineItem(dt_lineitem, "XBL", Convert.ToString(dt_LoadNo.Rows[0]["FieldValue"].ToString()), Convert.ToString(dt_LoadNo.Rows[0]["FieldColor"].ToString()));

                    }
                    else
                    {
                        //add in header field
                        dt_header = RemoveFieldRowFromDataTable(dt_header, cls_HeaderStaticField.Bol_No);
                        dt_header.Rows.Add("", cls_HeaderStaticField.Bol_No, Convert.ToString(dt_LoadNo.Rows[0]["FieldValue"].ToString()), Convert.ToString(dt_LoadNo.Rows[0]["FieldColor"].ToString()));

                    }
                }

            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in CheckLoadNoForBOL--->" + ex, LogType.E);
            }
            ds = CopyDatatableIntoMemory(dt_header, dt_lineitem);
            return ds;
        }
        public static DataSet SortMultiHeaderFieldValue(DataTable dt_SortedHeaderdata, DataTable dt_LineItem)
        {
            int Index = 0;
            int SrNO = 0;
            double RowNo = 0;
            int forCount = 0;
            DataSet ds = new DataSet();
            DataSet dsCheckLoadNoForBOL = new DataSet();
            List<cls_HeaderField> _ListHeaderField = new List<cls_HeaderField>();
            string[] _ArrayTillSplitterValue = null;
            string[] _ArrayTillsplitterColor = null;
            try
            {
                dsCheckLoadNoForBOL = CheckLoadNoForBOL(dt_SortedHeaderdata, dt_LineItem);
                dt_SortedHeaderdata = dsCheckLoadNoForBOL.Tables["Table1"];
                dt_LineItem = dsCheckLoadNoForBOL.Tables["Table2"];

                List<KeyValuePair<string, string>> _MultiValueFieldName = cls_AppConstant.MultiFieldValueFieldsName();
                _ListHeaderField = ConvertDataTable<cls_HeaderField>(dt_SortedHeaderdata);


                foreach (var itemValue in _MultiValueFieldName)
                {
                    try
                    {
                        var Headerdata = _ListHeaderField.Where(x => x.FieldName.ToLower().Trim() == itemValue.Key.ToLower().Trim()).FirstOrDefault();
                        if (Headerdata != null)
                        {
                            //r~u,1~1
                            //d~g,0~0
                            //PO~df

                            //***********Length >20 add into line item 25May21*********
                            string strBOL_PO_SN_Value = Headerdata.FieldValue;
                            string strBOL_PO_SN_Color = Headerdata.FieldColor;
                            if (strBOL_PO_SN_Value.Contains(cls_AppConstant.Till))
                            {
                                if (CheckFiledsValueLength(strBOL_PO_SN_Value.Substring(0, strBOL_PO_SN_Value.IndexOf('~')), cls_AppConstant.PO_BL_SP_length) == true)
                                {
                                    strBOL_PO_SN_Value = strBOL_PO_SN_Value.Substring(0, strBOL_PO_SN_Value.IndexOf('~')) + "~" + strBOL_PO_SN_Value;
                                    strBOL_PO_SN_Color= strBOL_PO_SN_Color.Substring(0, strBOL_PO_SN_Color.IndexOf('~')) + "~" + strBOL_PO_SN_Color;
                                }
                            }
                            else
                            {
                                if (CheckFiledsValueLength(strBOL_PO_SN_Value, cls_AppConstant.PO_BL_SP_length) == true)
                                {
                                    strBOL_PO_SN_Value = strBOL_PO_SN_Value + "~" + strBOL_PO_SN_Value;
                                    strBOL_PO_SN_Color = strBOL_PO_SN_Color + "~" + strBOL_PO_SN_Color;
                                }
                            }

                            

                            //************************************************************

                            if (strBOL_PO_SN_Value.Contains(cls_AppConstant.Till))
                            {
                                _ArrayTillSplitterValue = TillSplitter(strBOL_PO_SN_Value);
                                _ArrayTillsplitterColor = TillSplitter(strBOL_PO_SN_Color);

                                ApplyRules_BOL_PO_SN(_ArrayTillSplitterValue, _ArrayTillsplitterColor);

                            }
                            else
                            {
                                List<string> listfieldvalue = new List<string>();
                                List<string> listfieldcolor = new List<string>();
                                listfieldvalue.Add(strBOL_PO_SN_Value);
                                listfieldcolor.Add(strBOL_PO_SN_Color);
                                Arry_MultipleFieldValue = listfieldvalue.ToArray();
                                Array_MultiFieldColor = listfieldcolor.ToArray();

                            }

                            string AS400_Value = string.Empty;

                            //***********commanted 27April2021 since not working EDI Document*****
                            //AS400_Value = ReadPo_BL_SN_AS400(itemValue.Key);
                            //***************************************************************

                            if (!string.IsNullOrEmpty(AS400_Value))
                            {
                                Arry_MultipleFieldValue = CheckSameValueOnAS400(AS400_Value, Arry_MultipleFieldValue.ToList());
                                _ListHeaderField = _ListHeaderField.Where(x => x.FieldName.ToLower().Trim() != itemValue.Key.ToLower().Trim()).ToList();
                                forCount = 0;
                            }
                            else
                            {
                                _ListHeaderField.ForEach(x =>
                                {
                                    if (x.FieldName.ToLower().Trim() == itemValue.Key.ToLower().Trim())
                                    {
                                        x.FieldValue = Arry_MultipleFieldValue[0];//5
                                        x.FieldColor = Array_MultiFieldColor[0];
                                    }
                                });
                                forCount = 1;
                            }
                            if (dt_LineItem.Rows.Count > 0)
                            {
                                Index = dt_LineItem.Rows.Count - 1;
                                SrNO = Convert.ToInt32(dt_LineItem.Rows[Index]["SrNo"].ToString());
                                RowNo = Convert.ToDouble(dt_LineItem.Rows[Index]["RowNo"].ToString());
                            }
                            for (int i = forCount; i < Arry_MultipleFieldValue.Length; i++)//0
                            {

                                SrNO++;
                                RowNo++;
                                dt_LineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, itemValue.Value, Array_MultiFieldColor[i], RowNo });
                                SrNO++;
                                dt_LineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, Arry_MultipleFieldValue[i], Array_MultiFieldColor[i], RowNo });

                            }
                            //0

                            Arry_MultipleFieldValue = null;
                            Array_MultiFieldColor = null;

                        }
                    }
                    catch (Exception ex)
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception Message in _MultiValueFieldName foreach if condition--->" + ex,LogType.E );
                    }
                }

                if (_ListHeaderField.Count > 0)
                {
                    dt_SortedHeaderdata = ConvertToDataTable(_ListHeaderField);
                    dt_SortedHeaderdata.Columns.Remove("RowNo");
                    ds = CopyDatatableIntoMemory(dt_SortedHeaderdata, dt_LineItem);
                    return ds;

                }

            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception Message in SortMultiHeaderFieldValue--->" + ex,LogType.E );
            }
            ds = CopyDatatableIntoMemory(dt_SortedHeaderdata, dt_LineItem);
            return ds;
        }
        public static DataSet SortMultiHeaderFieldValue_Backup(DataTable dt_SortedHeaderdata, DataTable dt_LineItem)
        {
            int Index = 0;
            int SrNO = 0;
            double RowNo = 0;
            int forCount = 0;
            DataSet ds = new DataSet();
            DataSet dsCheckLoadNoForBOL = new DataSet();
            List<cls_HeaderField> _ListHeaderField = new List<cls_HeaderField>();
            string[] _ArrayTillSplitterValue = null;
            string[] _ArrayTillsplitterColor = null;
            try
            {
                dsCheckLoadNoForBOL = CheckLoadNoForBOL(dt_SortedHeaderdata, dt_LineItem);
                dt_SortedHeaderdata = dsCheckLoadNoForBOL.Tables["Table1"];
                dt_LineItem = dsCheckLoadNoForBOL.Tables["Table2"];

                List<KeyValuePair<string, string>> _MultiValueFieldName = cls_AppConstant.MultiFieldValueFieldsName();
                _ListHeaderField = ConvertDataTable<cls_HeaderField>(dt_SortedHeaderdata);


                foreach (var itemValue in _MultiValueFieldName)
                {
                    try
                    {
                        var Headerdata = _ListHeaderField.Where(x => x.FieldName.ToLower().Trim() == itemValue.Key.ToLower().Trim()).FirstOrDefault();
                        if (Headerdata != null)
                        {
                            //r~u,1~1
                            //d~g,0~0
                            //PO~df

                            //***********Length >20 add into line item 25May21*********

                            if (Headerdata.FieldValue.Contains(cls_AppConstant.Till))
                            {
                                _ArrayTillSplitterValue = TillSplitter(Headerdata.FieldValue);
                                _ArrayTillsplitterColor = TillSplitter(Headerdata.FieldColor);

                                ApplyRules_BOL_PO_SN(_ArrayTillSplitterValue, _ArrayTillsplitterColor);

                            }
                            else
                            {
                                List<string> listfieldvalue = new List<string>();
                                List<string> listfieldcolor = new List<string>();
                                listfieldvalue.Add(Headerdata.FieldValue);
                                listfieldcolor.Add(Headerdata.FieldColor);
                                Arry_MultipleFieldValue = listfieldvalue.ToArray();
                                Array_MultiFieldColor = listfieldcolor.ToArray();

                            }

                            string AS400_Value = string.Empty;

                            //***********commanted 27April2021 since not working EDI Document*****
                            //AS400_Value = ReadPo_BL_SN_AS400(itemValue.Key);
                            //***************************************************************

                            if (!string.IsNullOrEmpty(AS400_Value))
                            {
                                Arry_MultipleFieldValue = CheckSameValueOnAS400(AS400_Value, Arry_MultipleFieldValue.ToList());
                                _ListHeaderField = _ListHeaderField.Where(x => x.FieldName.ToLower().Trim() != itemValue.Key.ToLower().Trim()).ToList();
                                forCount = 0;
                            }
                            else
                            {
                                _ListHeaderField.ForEach(x =>
                                {
                                    if (x.FieldName.ToLower().Trim() == itemValue.Key.ToLower().Trim())
                                    {
                                        x.FieldValue = Arry_MultipleFieldValue[0];//5
                                        x.FieldColor = Array_MultiFieldColor[0];
                                    }
                                });
                                forCount = 1;
                            }
                            if (dt_LineItem.Rows.Count > 0)
                            {
                                Index = dt_LineItem.Rows.Count - 1;
                                SrNO = Convert.ToInt32(dt_LineItem.Rows[Index]["SrNo"].ToString());
                                RowNo = Convert.ToDouble(dt_LineItem.Rows[Index]["RowNo"].ToString());
                            }
                            for (int i = forCount; i < Arry_MultipleFieldValue.Length; i++)//0
                            {

                                SrNO++;
                                RowNo++;
                                dt_LineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, itemValue.Value, Array_MultiFieldColor[i], RowNo });
                                SrNO++;
                                dt_LineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, Arry_MultipleFieldValue[i], Array_MultiFieldColor[i], RowNo });

                            }
                            //0

                            Arry_MultipleFieldValue = null;
                            Array_MultiFieldColor = null;

                        }
                    }
                    catch (Exception ex)
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception Message in _MultiValueFieldName foreach if condition--->" + ex, LogType.E);
                    }
                }

                if (_ListHeaderField.Count > 0)
                {
                    dt_SortedHeaderdata = ConvertToDataTable(_ListHeaderField);
                    dt_SortedHeaderdata.Columns.Remove("RowNo");
                    ds = CopyDatatableIntoMemory(dt_SortedHeaderdata, dt_LineItem);
                    return ds;

                }

            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception Message in SortMultiHeaderFieldValue--->" + ex, LogType.E);
            }
            ds = CopyDatatableIntoMemory(dt_SortedHeaderdata, dt_LineItem);
            return ds;
        }

        public static string[] CheckSameValueOnAS400(string AS400_Value, List<string> _InputlistBL_PO_SN_Convert)
        {

            string[] As400CommaValueArray = null;
            List<string> _ListAddPO_BL_SN = new List<string>();//_InputlistBL_PO_SN_Convert;
            _ListAddPO_BL_SN.AddRange(_InputlistBL_PO_SN_Convert);
            try
            {


                if (!string.IsNullOrEmpty(AS400_Value))//000
                {
                    if (AS400_Value.Contains(cls_AppConstant.Comma))
                    {
                        As400CommaValueArray = AS400_Value.Split(',');
                        if (As400CommaValueArray.Length > 0)
                        {


                            for (int i = 0; i < _InputlistBL_PO_SN_Convert.Count; i++)
                            {
                                foreach (var item in As400CommaValueArray)//2
                                {
                                    if (_InputlistBL_PO_SN_Convert[i].ToLower().Trim() == item.ToLower().Trim())//match value with client UI value
                                    {
                                        _ListAddPO_BL_SN.RemoveAt(i);//
                                                                     //1

                                    }
                                }
                            }
                        }
                    }
                    else
                    {


                        for (int i = 0; i < _InputlistBL_PO_SN_Convert.Count; i++)
                        {
                            if (_InputlistBL_PO_SN_Convert[i].ToLower().Trim() == AS400_Value.ToLower().Trim())//match value with client UI value
                            {

                                //Arry_MultipleFieldValue = _listBL_PO_SN_Convert.Where((source, _index) => _index != i).ToArray();
                                _ListAddPO_BL_SN.RemoveAt(i);


                            }
                        }
                    }

                    Arry_MultipleFieldValue = _ListAddPO_BL_SN.ToArray();

                }


            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Match Same PO BL SN Number in CheckSameValueOnAS400 function--->" + ex,LogType.E);
            }

            return Arry_MultipleFieldValue;
        }

        public static string GetPo_BL_SN_FRP001(string FiledName)
        {
            string Po_BL_SN = cls_FieldRuleValidation.HeaderItemMappingValue(FiledName, "");
            try
            {
                if (!string.IsNullOrEmpty(Po_BL_SN))
                {
                    SelectPOBLSNFromFRP001(Po_BL_SN, "");
                }
            }
            catch (Exception ex)
            {

            }

            return "";
            }
        public static string ReadPo_BL_SN_AS400(string FiledName)
        {
            string AS400_Value = string.Empty;
            try
            {
                List<KeyValuePair<string, string>> _MultiValueFieldName = cls_AppConstant.MultiFieldValueRowColumnAS400();
                _MultiValueFieldName = _MultiValueFieldName.Where(x => x.Key.Trim().ToLower() == FiledName.Trim().ToLower()).ToList();

                string[] RowcolumnLength = _MultiValueFieldName[0].Value.Split(',');

                AS400_Value = GetPo_BL_SN_FRP001(FiledName);//"Read value from DB";//cls_Global.ReadStr(Convert.ToInt32(RowcolumnLength[0]), Convert.ToInt32(RowcolumnLength[1]), Convert.ToInt32(RowcolumnLength[2]), cls_SaiaBol.str_SessionName).Trim();
                if (AS400_Value.ToLower().Trim() == cls_AppConstant.PO_BL_SN_NS.ToLower().Trim() || AS400_Value.ToLower().Trim() == cls_AppConstant.PO_BL_SN_NA.ToLower().Trim())
                {
                    //cls_Global.func_InfoDelete(Convert.ToInt32(RowcolumnLength[0]), Convert.ToInt32(RowcolumnLength[1]), Convert.ToInt32(RowcolumnLength[2]), cls_SaiaBol.str_SessionName);
                    AS400_Value = string.Empty;
                }
                return AS400_Value;

            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in ReadPo_BL_SN_AS400()--->" + ex, LogType.E);
                //cls_Global.func_ErrorLog("in ReadPo_BL_SN_AS400" + ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// Split value using till(~)
        /// </summary>
        /// <param name="Value"></param>
        /// <returns></returns>
        public static string[] TillSplitter(string Value)
        {
            string[] Array_TillSeperation = null;
            try
            {
                if (!string.IsNullOrEmpty(Value))
                {
                    if (Value.Contains(cls_AppConstant.Till))
                    {
                        Array_TillSeperation = Value.Split(cls_AppConstant.Till);
                        return Array_TillSeperation;
                    }
                }
                return Array_TillSeperation;
            }
            catch (Exception)
            {
                return Array_TillSeperation;
            }

        }

        /// <summary>
        /// find character in array string value
        /// and return value index
        /// method in use
        /// </summary>
        /// <param name="_ArrayValue"></param>
        /// <returns></returns>
        public static int CheckCharactersInValue(string[] _ArrayValue)
        {
            // string _FiledValue = string.Empty;
            int index = -1;
            try
            {
                if (_ArrayValue != null && _ArrayValue.Length > 0)
                {
                    foreach (var item in _ArrayValue)
                    {
                        index++;
                        char[] _chhar = item.ToCharArray();
                        foreach (char c in _chhar)
                        {
                            if (char.IsLetter(c))
                            {
                                return index;
                            }
                        }
                    }
                }
            }
            catch
            {

                return index;
            }
            return index;
        }


        /// <summary>
        /// if PO#,BL#,SN charachter find with value in perticular filed
        ///  return those field value index
        /// </summary>
        /// <param name="_ArrayValue"></param>
        /// <returns></returns>
        public static int CheckCharactersValueIndex(string[] _ArrayValue)
        {

            int index = -1;
            List<string> listitems = cls_AppConstant.StaticFieldNames.ToList();
            try
            {
                if (_ArrayValue != null)
                {
                    for (int i = 0; i < listitems.Count; i++)
                    {
                        index = Array.IndexOf(_ArrayValue, listitems[i]);
                        if (index > -1)
                        {
                            return index;
                        }
                    }
                }
            }
            catch
            {

                return index;
            }
            return index;
        }

        /// <summary>
        /// check count greater than 10 for BL,PO,SN
        /// </summary>
        /// <param name="_Array"></param>
        /// <returns></returns>
        public static bool CheckLength(string[] _Array)
        {
            if (_Array != null)
            {
                if (_Array.Length > 10)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
        }

        /// <summary>
        /// find index of charcter for PO,BL,SN
        /// if char find then set this value as primary field
        ///only take first 10 value in array
        ///check length of array 
        /// </summary>
        /// <param name="_ArrayFieldValue"></param>
        /// <param name="_ArrayFieldColor"></param>
        public static void ApplyRules_BOL_PO_SN(string[] _ArrayFieldValue, string[] _ArrayFieldColor)
        {
            try
            {

                if (_ArrayFieldValue != null && _ArrayFieldColor != null)
                {
                    //check value in record   attached with charachters
                    //remove characters from value
                    //assign character attached value in primary field
                    //if lenght greter than 10 then take only first 10 record  
                    //other 9 values assign in line items fields
                    //int index = CheckCharactersValueIndex(_ArrayFieldValue);

                    List<string> _FiledValue = _ArrayFieldValue.ToList();
                    List<string> _FieldColor = _ArrayFieldColor.ToList();

                    for (int i = 0; i < _ArrayFieldColor.Length; i++)
                    {
                        if (_ArrayFieldColor[i] == "0")
                        {
                            _FiledValue.RemoveAt(i);
                            _FieldColor.RemoveAt(i);
                        }
                    }


                    if (true == CheckLength(_FiledValue.ToArray()))
                    {

                        Arry_MultipleFieldValue = _FiledValue.Take(cls_AppConstant.RecordLength).Select(i => i.ToString()).ToArray();
                        Array_MultiFieldColor = _FieldColor.Take(cls_AppConstant.RecordLength).Select(i => i.ToString()).ToArray();
                    }
                    else
                    {
                        Arry_MultipleFieldValue = _FiledValue.ToArray();
                        Array_MultiFieldColor = _FieldColor.ToArray();
                    }

                    // Arry_MultipleFieldValue = RemoveColonsSemiColons(Arry_MultipleFieldValue);





                }

            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception ApplyRules_BOL_PO_SN" + ex,LogType.E );
                throw;
            }

        }

        public static string[] RemoveColonsSemiColons(string[] _ArrayFieldValue)
        {
            string fieldValue = string.Empty;

            List<string> _List = new List<string>();
            try
            {
                if (_ArrayFieldValue != null)
                {
                    _List = _ArrayFieldValue.ToList();
                    _List.ForEach(x =>
                    {
                        if (x.Contains(cls_AppConstant.Colons) || x.Contains(cls_AppConstant.SemiColons))
                        {
                            x = x.Replace(cls_AppConstant.Colons, "");
                            x = x.Replace(cls_AppConstant.SemiColons, "");
                        }
                    });
                    return _List.ToArray();
                }
            }
            catch (Exception)
            {
                return _List.ToArray();
            }
            return _List.ToArray();
        }

        public static string RemoveCharFromValue(string _ArrayFieldValue)
        {

            try
            {
                if (!string.IsNullOrEmpty(_ArrayFieldValue))
                {
                    foreach (var item in cls_AppConstant.StaticFieldNames)
                    {

                        if (_ArrayFieldValue.Contains(item))
                        {
                            _ArrayFieldValue = _ArrayFieldValue.Replace(item, "");
                            return _ArrayFieldValue;
                        }

                    }
                }
            }
            catch (Exception)
            {
                return string.Empty;

            }
            return string.Empty;
        }
        //20
        public static int CheckLength(string PO_Value)
        {
            if (!string.IsNullOrEmpty(PO_Value))
            {
                int length = PO_Value.Length;
                return length;
            }
            return -1;
        }


        /// <summary>
        /// this function use for check PO value length
        /// if PO Value is to long then do not other character in PO Department
        /// </summary>
        /// <param name="PO_Value"></param>
        /// <returns></returns>
        public static string GetPO_ValueUsingLength(string PO_Value)
        {
            try
            {
                if (!string.IsNullOrEmpty(PO_Value))
                {
                    int length = CheckLength(PO_Value);
                    if (length <= 20)
                    {
                        return PO_Value;
                    }
                    else
                    {
                        PO_Value = new string(PO_Value.Take(length).ToArray());
                        return PO_Value;
                    }
                }
                return string.Empty;
            }
            catch (Exception)
            {
                return string.Empty;

            }
        }




        #endregion


        #region Shiper/Cons/Third Party

        public static List<KeyValuePair<String, String>> MatchAccountCode(List<KeyValuePair<String, String>> ListFilter)
        {
            List<KeyValuePair<String, String>> _ListShipCons =new List<KeyValuePair<String, String>>();
            string ConsiValue = string.Empty;
            string ShipValue = string.Empty;
            try
            {
               var  _ListBillTo = ListFilter.Where(x => x.Key.ToLower().Trim() == cls_HeaderStaticField.Bill_To_Code.ToLower()).FirstOrDefault();
                
                if(!string.IsNullOrEmpty(_ListBillTo.Key) && !string.IsNullOrEmpty(_ListBillTo.Value))
                {
                    var _Listship = ListFilter.Where(x => x.Key.ToLower().Trim() == cls_HeaderStaticField.Shp_Code.ToLower()).FirstOrDefault();
                    var _ListCons = ListFilter.Where(x => x.Key.ToLower().Trim() == cls_HeaderStaticField.Con_Code.ToLower()).FirstOrDefault();
                    
                    if(_ListCons.Value!=null)
                    {
                        ConsiValue = Convert.ToString(_ListCons.Value).Trim();
                    }
                    if (_Listship.Value != null)
                    {
                        ShipValue = Convert.ToString(_Listship.Value).Trim();
                    }


                    if (Convert.ToString(_ListBillTo.Value).Trim()== ConsiValue || Convert.ToString(_ListBillTo.Value).Trim() == ShipValue)
                    {
                        ListFilter = ListFilter.Where(x => x.Key.ToLower().Trim() != cls_HeaderStaticField.Bill_To_Code.ToLower())
                            .Where(x => x.Key.ToLower().Trim() != cls_HeaderStaticField.Bill_To_Addr_Line.ToLower())
                            .Where(x => x.Key.ToLower().Trim() != cls_HeaderStaticField.Bill_To_City.ToLower())
                            .Where(x => x.Key.ToLower().Trim() != cls_HeaderStaticField.Bill_To_ZipCode.ToLower())
                            .Where(x => x.Key.ToLower().Trim() != cls_HeaderStaticField.Bill_To_State.ToLower())
                            .Where(x => x.Key.ToLower().Trim() != cls_HeaderStaticField.Bill_To_Name.ToLower())
                            .ToList();
                        
                        return ListFilter; 
                    }
                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in MatchAccountCode--->" + ex,LogType.E);
            }
            return ListFilter;
        }

        #endregion

        #region CodeName Seperate # in use

        //split using # and get only values
        public static string[] HashSplitterForCode(string _Hashvalue)
        {
            string[] HashCodename = null;
            try
            {
                if (!string.IsNullOrEmpty(_Hashvalue))
                {
                    if (_Hashvalue.Contains(cls_AppConstant.Hash))
                    {
                        HashCodename = _Hashvalue.Split(cls_AppConstant.HashSplit);
                        return HashCodename;
                    }
                   // else { return HashCodename; }
                }
            }
            catch (Exception)
            {
                return HashCodename;
                throw;
            }
            return HashCodename;
        }
        #endregion

        #region COD in use

        /// <summary>
        /// RemoveDot signfrom string for validation
        /// </summary>
        /// <param name="_Value"></param>
        /// <returns></returns>


        /// <summary>
        /// check COD filed cashpayment or not
        /// if COD value is not null then payment is in cash
        /// if COD then add value in lineitems
        /// 1.FEE in code and other value in Description
        /// 2.CASH in code other value in description
        /// </summary>
        /// <param name="ds"></param>
        /// <returns></returns>
        public static DataSet CODCashPaymentOrNot(DataSet ds)
        {
            DataSet ds_COD = new DataSet();
            DataTable dtHeader = new DataTable();
            DataTable dtLineItem = new DataTable();
            bool IsCOD = false;

            try
            {
                if (ds.Tables.Count > 0)
                {
                    dtHeader = ds.Tables["Table1"];
                    dtLineItem = ds.Tables["Table2"];

                    DataTable dt_COD = FindFieldRowInDataTable(dtHeader, cls_AppConstant.COD);
                    dtHeader = RemoveFieldRowFromDataTable(dtHeader, cls_AppConstant.COD);
                    if (dt_COD.Rows.Count > 0)
                    {
                        string CODValue = dt_COD.Rows[0]["FieldValue"].ToString();
                        if (!string.IsNullOrEmpty(CODValue))
                        {
                            IsCOD = CheckCODInvalidValue(CODValue);
                            if (IsCOD)
                            {
                                //
                                CODValue = RemoveDot(CODValue);
                                CODValue = RemoveDollerSign(CODValue);
                                //for cash payment add value line items
                                //if (dtLineItem.Rows.Count > 0)
                                //{
                                //    Index = dtLineItem.Rows.Count - 1;
                                //    SrNO = Convert.ToInt32(dtLineItem.Rows[Index]["SrNo"].ToString());
                                //    RowNo = Convert.ToDouble(dtLineItem.Rows[Index]["RowNo"].ToString());
                                //}
                                List<KeyValuePair<string, string>> CODCASH = cls_AppConstant.CODCashArray();
                                foreach (var item in CODCASH)
                                {
                                    dtLineItem = AddFieldValueinLineItem(dtLineItem, item.Key, item.Value, dt_COD.Rows[0]["FieldColor"].ToString());
                                    //SrNO++;
                                    //RowNo++;
                                    //dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, item.Key, dt_COD.Rows[0]["FieldColor"].ToString(), RowNo });
                                    //SrNO++;
                                    //dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, item.Value, dt_COD.Rows[0]["FieldColor"].ToString(), RowNo });

                                }
                                int Index = dtHeader.Rows.Count - 1;
                                int SrNO = Convert.ToInt32(dtHeader.Rows[Index]["SrNo"].ToString());
                                dtHeader.Rows.Add(SrNO, cls_AppConstant.COD, CODValue, dt_COD.Rows[0]["FieldColor"].ToString());


                            }

                            ds_COD = CopyDatatableIntoMemory(dtHeader, dtLineItem);
                            return ds_COD;
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-exception in CODCashPaymentOrNot As400--->" + ex,LogType.E);
                return ds;

            }
            return ds;
        }

        public static bool CheckCODInvalidValue(string _CODValue)
        {
            try
            {
                if (!string.IsNullOrEmpty(_CODValue))
                {
                    foreach (var item in cls_AppConstant.CODInvalidValue)
                    {

                        //0.00=0.00
                        if (_CODValue.Trim() == item.Trim())
                        {
                            return false;
                        }
                    }


                }
                return true;
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-exception in CheckCODInvalidValue As400--->" + ex, LogType.E);
                return false;
            }
        }
        #endregion



       

        #region For Address

        public static List<string> FindAndSetAddress(string AddressValue)
        {
            List<string> _ListAddress = new List<string>();
            try
            {
                if (!string.IsNullOrEmpty(AddressValue))
                {
                    _ListAddress = cls_FieldRuleValidation.SortAddressLine(AddressValue);

                    if (_ListAddress.Count > 0)
                    {

                        return _ListAddress;
                    }
                    else
                    {
                        List<string> _listShipaddress = cls_FieldRuleValidation.SetTextOnNewLine(AddressValue, cls_AppConstant.AddressLength);

                        return _listShipaddress;
                    }
                }
                return _ListAddress;

            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("exception in FindAndSetAddress function--->" + ex,LogType.E );
                return _ListAddress;
            }

        }
        public static List<string> SortAddressLine(string AddressValue)
        {

            List<string> _listAddress = new List<string>();
            try
            {
                if (AddressValue.Contains(cls_AppConstant.DoublePipe))
                {
                    string[] _AddressArray = AddressValue.Split(new[] { cls_AppConstant.DoublePipe }, StringSplitOptions.None);
                    if (_AddressArray != null && _AddressArray.Length > 0)
                    {
                        _listAddress = _AddressArray.ToList();
                    }
                }


            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("exception in SortAddressLine function-->" + ex,LogType.E );

            }
            return _listAddress;
        }

        #endregion


        #region common methods

        public static string TrimFieldValueLength(string FieldValue, int TrimLength)
        {
            try
            {
                if (!string.IsNullOrEmpty(FieldValue))
                {
                    FieldValue = FieldValue.Trim();

                    if (FieldValue.Length > TrimLength)
                    {
                        FieldValue = FieldValue.Substring(0, TrimLength).Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("exception in TrimFieldValueLength function--->" + ex, LogType.E);
                // return _ListAddress;
            }

            return FieldValue;

        }

        public static bool CheckFiledsValueLength(string FieldValue, int TrimLength)
        {
            bool ReturnType = false;
            try
            {
                if (!string.IsNullOrEmpty(FieldValue))
                {
                    FieldValue = FieldValue.Trim();

                    if (FieldValue.Length > TrimLength)
                    {
                        ReturnType = true;
                    }
                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("exception in CheckFiledsValueLength function--->" + ex, LogType.E);
                // return _ListAddress;
            }

            return ReturnType;

        }

        public static string Class_CodeValidate(string _ClassCode)
        {
            try
            {
                string[] _Array = null;

                if (!string.IsNullOrEmpty(_ClassCode))
                {

                    foreach (var item in cls_AppConstant.ClassCodeArray)
                    {
                        if (_ClassCode.Trim().Contains(item))
                        {
                            _Array = _ClassCode.Split('.');
                            _ClassCode =Convert.ToString( _Array[0]);
                            return _ClassCode;
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in Class_CodeValidate function--->" + ex,LogType.E );
            }
            return _ClassCode;
        }

        /// <summary>
        ///
        /// manage text with textbx length in line items
        /// if text is too long as compate length of textbox than remaining text set in new line
        /// count character and length
        /// </summary>
        /// <param name="DescriptionValue"></param>
        /// <returns></returns>
        public static List<string> SetTextOnNewLine(string _TextValue, int _TextLength)
        {
            string _StringTextValue = string.Empty;
            List<string> _ListSetText = new List<string>();

            try
            {
                if (!string.IsNullOrEmpty(_TextValue))
                {
                    if (_TextValue.Length > _TextLength)
                    {
                        _TextValue = FindAndReplacePunctuation(_TextValue);
                        // sagar vivek sushant,sandeep

                        string[] _SplitDescArray = _TextValue.Split(' ');
                        foreach (var item in _SplitDescArray)
                        {
                            int matchlen = (_StringTextValue.Length + item.Length);
                            if (matchlen < _TextLength)//4 > 35
                            {
                                string _value = item.Trim();
                                _StringTextValue += " " + _value.ToString().Trim();
                            }
                            else
                            {
                                _ListSetText.Add(_StringTextValue);
                                _StringTextValue = string.Empty;
                                _StringTextValue = item;
                            }
                        }
                        _ListSetText.Add(_StringTextValue);
                    }
                    else
                    {
                        _ListSetText.Add(_TextValue);
                        return _ListSetText;
                    }
                }
                else
                {
                    _ListSetText.Add(_TextValue);
                    return _ListSetText;
                }

            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in SetTextOnNewLine function--->" + ex,LogType.E );
                return _ListSetText;
            }
            return _ListSetText;
        }


        public static string FindAndReplacePunctuation(string _value)
        {
            try
            {
                if (!string.IsNullOrEmpty(_value))
                {
                    foreach (var item in cls_AppConstant.Punctuation)
                    {
                        if (_value.Contains(item))
                        {
                            _value = _value.Replace(item, item + cls_AppConstant.Replacewithspace);
                        }
                    }
                    return _value;

                }
            }
            catch (Exception)
            {
                return _value;
                throw;
            }
            return _value;
        }
        public static DataTable FindFieldRowInDataTable(DataTable dt_Input, string FieldName)
        {
            DataTable dt = new DataTable();
            try
            {

                var dtVar = dt_Input.AsEnumerable().Where(r => r.Field<string>("FieldName").ToString().Trim().ToLower() == FieldName.ToLower().Trim()).ToList();
                if (dtVar.Any())
                {
                    dt = dtVar.CopyToDataTable();
                }
            }
            catch (Exception ex)
            {

                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in FindFieldRowInDataTable common method region-->" + ex,LogType.E );
            }
            return dt;

        }

        public static DataTable RemoveFieldRowFromDataTable(DataTable dt_Input, string FieldName)
        {
            DataTable dt = new DataTable();
            try
            {
                var dtVar = dt_Input.AsEnumerable().Where(r => r.Field<string>("FieldName").ToString().Trim().ToLower() != FieldName.ToLower().Trim()).ToList();
                if (dtVar.Any())
                {
                    dt = dtVar.CopyToDataTable();
                }
            }
            catch (Exception ex)
            {

                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in RemoveFieldRowFromDataTable common method region-->" + ex,LogType.E );
            }
            return dt;

        }

        public static DataTable AddFieldValueinLineItem(DataTable dtLineItem, string Class_CodeValue, string DescriptionValue, string FieldColor)
        {
            int Index = 0;
            int SrNO = 0;
            double RowNo = 0;

            try
            {
                if (dtLineItem.Rows.Count > 0)
                {
                    Index = dtLineItem.Rows.Count - 1;
                    SrNO = Convert.ToInt32(dtLineItem.Rows[Index]["SrNo"].ToString());
                    RowNo = Convert.ToDouble(dtLineItem.Rows[Index]["RowNo"].ToString());
                }
                SrNO++;
                RowNo++;
                dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, Class_CodeValue, FieldColor, RowNo });
                SrNO++;
                dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, DescriptionValue, FieldColor, RowNo });
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in AddExtraFieldValueinLineItem--->" + ex,LogType.E);
            }

            return dtLineItem;
        }

        /// <summary>
        /// check lineItem filevalue chec isnullor empty
        /// </summary>
        /// <param name="_Value"></param>
        /// <returns></returns>
        public static string CheckIsnullOrEmpty(string _Value)
        {
            try
            {
                if (string.IsNullOrEmpty(_Value))
                {
                    _Value = string.Empty;
                    return _Value;
                }
                else
                {
                    return _Value.Trim();
                }
            }
            catch (Exception)
            {
                return _Value;
            }

        }

        public static string RemoveDot(string _Value)
        {
            try
            {
                if (!string.IsNullOrEmpty(_Value))
                {
                    if (_Value.Contains(cls_AppConstant.DotSign))
                    {
                        _Value = _Value.Replace(cls_AppConstant.DotSign, "");
                        return _Value;
                    }
                }
                return _Value;
            }
            catch (Exception)
            {
                return _Value;

            }

        }

        public static DataSet RemoveStarFieldValue(DataTable dtheader, DataTable dtLineITem)
        {
            DataSet ds = new DataSet();


            try
            {
                if (dtheader.Rows.Count > 0)
                {
                    for (int i = 0; i < dtheader.Rows.Count; i++)
                    {
                        string _FiledValue = dtheader.Rows[i]["FieldValue"].ToString();
                        if (_FiledValue.Contains("**"))
                        {

                            dtheader.Rows[i]["FieldValue"] = string.Empty;
                            // dtheader = dtheader.AsEnumerable().Where(x => x.Field<string>("FieldValue") != _FiledValue).CopyToDataTable();

                        }
                    }
                }
                if (dtheader.Rows.Count > 0)
                {
                    for (int i = 0; i < dtLineITem.Rows.Count; i++)
                    {
                        string _FiledValue = dtLineITem.Rows[i]["FieldValue"].ToString();
                        if (_FiledValue.Contains("**"))
                        {
                            dtLineITem.Rows[i]["FieldValue"] = string.Empty;
                            //dtLineITem = dtLineITem.AsEnumerable().Where(x => x.Field<string>("FieldValue") != _FiledValue).CopyToDataTable();
                        }
                    }
                }
                ds = CopyDatatableIntoMemory(dtheader, dtLineITem);
            }
            catch (Exception ex)
            {

                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in RemoveStar--->" + ex,LogType.E);
            }
            return ds;

        }
        public static string RemoveDollerSign(string _Value)
        {
            try
            {
                if (!string.IsNullOrEmpty(_Value))
                {
                    _Value = _Value.Replace(cls_AppConstant.DollerSign, "");
                    return _Value;
                }
            }
            catch (Exception)
            {
                return _Value;

            }
            return _Value;
        }

        public static DataSet CopyDatatableIntoMemory(DataTable dtHeader, DataTable dtLineItem)
        {
            DataSet ds = new DataSet();
            DataTable dtCopyHeader = new DataTable();
            DataTable dtCopyLineItem = new DataTable();
            try
            {
                dtCopyHeader = dtHeader.Copy();
                dtCopyLineItem = dtLineItem.Copy();
                ds.Tables.Add(dtCopyHeader);
                ds.Tables.Add(dtCopyLineItem);
                dtCopyHeader = null;
                dtCopyLineItem = null;
                dtHeader = null;
                dtLineItem = null;
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception add DataTable in Dataset CopyDatatableIntoMemory--->" + ex,LogType.E );

            }
            return ds;
        }
        #endregion

        #region LineItems in use


        #region  Add another field in LineItem with Code Special_Instruction,NFMC  
        public static DataTable FindFieldNameInDataTable(DataTable dt_Input, string FieldName)
        {
            DataTable dt = new DataTable();
            try
            {

                var dtVar = dt_Input.AsEnumerable()
                    .Where(r => r.Field<string>("FieldName").ToString().Trim().ToLower() == FieldName.ToLower().Trim()).ToList();
                // && r.Field<string>("FieldColor").Trim() == cls_AppConstant.FieldColorOne).ToList();
                if (dtVar.Any())
                {
                    dt = dtVar.CopyToDataTable();
                }
            }
            catch (Exception ex)
            {

                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in FindFieldRowInDataTable common method region-->" + ex,LogType.E );
            }
            return dt;

        }

        public static DataTable CheckFiledValueForDescription(DataTable dt_Input, string FieldName)
        {
            DataTable dt = new DataTable();
            try
            {

                var dtVar = dt_Input.AsEnumerable()
                    .Where(r => r.Field<string>("FieldName").ToString().Trim().ToLower() == FieldName.ToLower().Trim()).ToList();
                if (dtVar.Any())
                {
                    dt = dtVar.CopyToDataTable();
                }
            }
            catch (Exception ex)
            {

                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in FindFieldRowInDataTable common method region-->" + ex,LogType.E );
            }
            return dt;

        }

        public static DataSet FindDescription(DataSet ds)
        {
            DataTable dt_lineItem = new DataTable();
            DataTable dt_Desc = new DataTable();
            DataTable _dtFileterRowNo = new DataTable();
            DataSet dsCopy = new DataSet();
            DataTable _NewLineitem = new DataTable();

            int _RorNOCounter = 1;

            _NewLineitem.Columns.Add("SrNo");
            _NewLineitem.Columns.Add("FieldName");
            _NewLineitem.Columns.Add("FieldValue");
            _NewLineitem.Columns.Add("FieldColor");
            _NewLineitem.Columns.Add("RowNo", typeof(double));
            try
            {
                dt_lineItem = ds.Tables["Table2"];
                dt_Desc = FindFieldNameInDataTable(dt_lineItem, LineItemStaticField.Description);
                int startcount = 1;
                if (dt_Desc.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_Desc.Rows.Count; i++)
                    {
                        string Rowno = Convert.ToString(dt_Desc.Rows[i]["RowNo"]);
                        var result = dt_lineItem.AsEnumerable().Where(x => x.Field<double>("RowNo") == Convert.ToDouble(Rowno)).ToList();


                        if (result.Any())
                        {

                            _dtFileterRowNo = result.CopyToDataTable();
                            string _FieldValue = Convert.ToString(dt_Desc.Rows[i]["FieldValue"]);

                            startcount = _RorNOCounter;
                            for (int j = 0; j < _dtFileterRowNo.Rows.Count; j++)
                            {

                                string _DescriptionName = Convert.ToString(_dtFileterRowNo.Rows[j]["FieldName"]);
                                string _DescriptionValue = Convert.ToString(_dtFileterRowNo.Rows[j]["FieldValue"]);
                                string _FieldColor = Convert.ToString(_dtFileterRowNo.Rows[j]["FieldColor"]);
                                string _SrNo = Convert.ToString(_dtFileterRowNo.Rows[j]["SrNo"]);
                                string _RowNo = string.Empty;
                                _RowNo = Convert.ToString(_dtFileterRowNo.Rows[j]["RowNo"]);

                                if (_DescriptionName == LineItemStaticField.Description && Rowno == _RowNo)
                                {
                                    List<string> _ListNMFCValues = cls_FieldRuleValidation.SetTextOnNewLine(_DescriptionValue, cls_AppConstant.DescriptionTextLength);

                                    if (_ListNMFCValues.Count > 0)
                                    {
                                        if (_ListNMFCValues.Count > 1)
                                        {

                                            //2 sagar vivek
                                            _NewLineitem.Rows.Add(new object[] { _SrNo, _DescriptionName, _ListNMFCValues[0], _FieldColor, _RorNOCounter });

                                            _RorNOCounter++;
                                            //class code
                                            _NewLineitem.Rows.Add(new object[] { _SrNo, "Class_Code", "", _FieldColor, _RorNOCounter });

                                            //description
                                            _NewLineitem.Rows.Add(new object[] { _SrNo, _DescriptionName, _ListNMFCValues[1], _FieldColor, _RorNOCounter });


                                        }
                                        else
                                        {

                                            _NewLineitem.Rows.Add(new object[] { _SrNo, _DescriptionName, _DescriptionValue, _FieldColor, _RorNOCounter });
                                        }
                                        _RorNOCounter++;
                                    }


                                }
                                else
                                {

                                    _NewLineitem.Rows.Add(new object[] { _SrNo, _DescriptionName, _DescriptionValue, _FieldColor, startcount });

                                }
                            }
                        }
                    }

                    dsCopy = CopyDatatableIntoMemory(ds.Tables["Table1"], _NewLineitem);
                }
                else
                {
                    dsCopy = CopyDatatableIntoMemory(ds.Tables["Table1"], dt_lineItem);
                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in append NMFC value in line item--->" + ex,LogType.E );
            }
            return dsCopy;

        }

        /// <summary>
        /// /Append NMFC lineitem column value in  in description
        /// </summary>
        /// <param name="ds"></param>
        /// <returns></returns>
        public static DataSet NMFC_old(DataSet ds)
        {
            DataTable dt_lineItem = new DataTable();
            DataTable dt_NFMC = new DataTable();
            DataTable _dtFileterRowNo = new DataTable();
            DataSet dsCopy = new DataSet();
            DataTable _NewLineitem = new DataTable();

            int _RorNOCounter = 1;

            _NewLineitem.Columns.Add("SrNo");
            _NewLineitem.Columns.Add("FieldName");
            _NewLineitem.Columns.Add("FieldValue");
            _NewLineitem.Columns.Add("FieldColor");
            _NewLineitem.Columns.Add("RowNo", typeof(double));
            try
            {
                dt_lineItem = ds.Tables["Table2"];
                dt_NFMC = FindFieldNameInDataTable(dt_lineItem, LineItemStaticField.NMFC);
                int startcount = 1;
                if (dt_NFMC.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_NFMC.Rows.Count; i++)
                    {
                        string Rowno = Convert.ToString(dt_NFMC.Rows[i]["RowNo"]);
                        var result = dt_lineItem.AsEnumerable().Where(x => x.Field<double>("RowNo") == Convert.ToDouble(Rowno)).ToList();


                        if (result.Any())
                        {

                            _dtFileterRowNo = result.CopyToDataTable();
                            string _NMFCFieldValue = Convert.ToString(dt_NFMC.Rows[i]["FieldValue"]);

                            startcount = _RorNOCounter;
                            for (int j = 0; j < _dtFileterRowNo.Rows.Count; j++)
                            {

                                string _DescriptionName = Convert.ToString(_dtFileterRowNo.Rows[j]["FieldName"]);
                                string _DescriptionValue = Convert.ToString(_dtFileterRowNo.Rows[j]["FieldValue"]);
                                string _FieldColor = Convert.ToString(_dtFileterRowNo.Rows[j]["FieldColor"]);
                                string _SrNo = Convert.ToString(_dtFileterRowNo.Rows[j]["SrNo"]);
                                string _RowNo = string.Empty;
                                _RowNo = Convert.ToString(_dtFileterRowNo.Rows[j]["RowNo"]);

                                if (_DescriptionName == LineItemStaticField.Description && Rowno == _RowNo)
                                {
                                    List<string> _ListNMFCValues = cls_FieldRuleValidation.SetTextOnNewLine(_DescriptionValue + "-" + _NMFCFieldValue, cls_AppConstant.DescriptionTextLength);

                                    if (_ListNMFCValues.Count > 0)
                                    {
                                        if (_ListNMFCValues.Count > 1)
                                        {

                                            //2 sagar vivek
                                            _NewLineitem.Rows.Add(new object[] { _SrNo, _DescriptionName, _ListNMFCValues[0], cls_AppConstant.FieldColorOne, _RorNOCounter });

                                            _RorNOCounter++;
                                            //class code
                                            _NewLineitem.Rows.Add(new object[] { _SrNo, "Class_Code", "", cls_AppConstant.FieldColorOne, _RorNOCounter });

                                            //description
                                            _NewLineitem.Rows.Add(new object[] { _SrNo, _DescriptionName, _ListNMFCValues[1], cls_AppConstant.FieldColorOne, _RorNOCounter });
                                            // _NewLineitem = AddFieldValueinLineItem(_NewLineitem, "", _ListNMFCValues[1], cls_AppConstant.FieldColorOne);

                                        }
                                        else
                                        {
                                            _DescriptionValue = _DescriptionValue + "-" + _NMFCFieldValue;
                                            _NewLineitem.Rows.Add(new object[] { _SrNo, _DescriptionName, _DescriptionValue, cls_AppConstant.FieldColorOne, _RorNOCounter });
                                        }
                                        _RorNOCounter++;
                                    }
                                }
                                else
                                {

                                    _NewLineitem.Rows.Add(new object[] { _SrNo, _DescriptionName, _DescriptionValue, _FieldColor, startcount });

                                }
                            }
                        }
                    }

                    dsCopy = CopyDatatableIntoMemory(ds.Tables["Table1"], _NewLineitem);
                }
                else
                {
                    dsCopy = CopyDatatableIntoMemory(ds.Tables["Table1"], dt_lineItem);
                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in append NMFC value in line item--->" + ex,LogType.E);
            }
            return dsCopy;

        }


        public static DataSet NMFC(DataTable dt_Headerdata, DataTable dt_lineItem)
        {

            DataTable dt_NFMC = new DataTable();
            DataTable _dtFileterRowNo = new DataTable();
            DataSet dsCopy = new DataSet();
            DataTable _NewLineitem = new DataTable();

            int _RorNOCounter = 1;

            _NewLineitem.Columns.Add("SrNo");
            _NewLineitem.Columns.Add("FieldName");
            _NewLineitem.Columns.Add("FieldValue");
            _NewLineitem.Columns.Add("FieldColor");
            _NewLineitem.Columns.Add("RowNo", typeof(double));
            try
            {

                dt_NFMC = FindFieldNameInDataTable(dt_lineItem, LineItemStaticField.NMFC);
                int startcount = 1;
                if (dt_NFMC.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_NFMC.Rows.Count; i++)
                    {
                        string Rowno = Convert.ToString(dt_NFMC.Rows[i]["RowNo"]);
                        var result = dt_lineItem.AsEnumerable().Where(x => x.Field<double>("RowNo") == Convert.ToDouble(Rowno)).ToList();


                        if (result.Any())
                        {

                            _dtFileterRowNo = result.CopyToDataTable();
                            string _NMFCFieldValue = Convert.ToString(dt_NFMC.Rows[i]["FieldValue"]);
                            string _NMFCFieldColor = Convert.ToString(dt_NFMC.Rows[i]["FieldColor"]);

                            startcount = _RorNOCounter;
                            for (int j = 0; j < _dtFileterRowNo.Rows.Count; j++)
                            {

                                string _DescriptionName = Convert.ToString(_dtFileterRowNo.Rows[j]["FieldName"]);
                                string _DescriptionValue = Convert.ToString(_dtFileterRowNo.Rows[j]["FieldValue"]);
                                string _FieldColor = Convert.ToString(_dtFileterRowNo.Rows[j]["FieldColor"]);
                                string _SrNo = Convert.ToString(_dtFileterRowNo.Rows[j]["SrNo"]);
                                string _RowNo = string.Empty;
                                _RowNo = Convert.ToString(_dtFileterRowNo.Rows[j]["RowNo"]);

                                if (_DescriptionName == LineItemStaticField.Description && Rowno == _RowNo)
                                {
                                    //17 march 2021--add check compare NFMC fieldcolor equal to Desscription FieldColor
                                    //if NFMC Color and Description color match then concat fieldvalue
                                    //if do not match then as it is pass fieldvalue to next stage
                                    if (_NMFCFieldColor.Trim() == cls_AppConstant.FieldColorOne && _FieldColor.Trim() == cls_AppConstant.FieldColorOne)
                                    {
                                        _DescriptionValue = _DescriptionValue + "-" + _NMFCFieldValue;
                                        _NewLineitem.Rows.Add(new object[] { _SrNo, _DescriptionName, _DescriptionValue, cls_AppConstant.FieldColorOne, _RowNo });

                                    }
                                    else
                                    {
                                        _NewLineitem.Rows.Add(new object[] { _SrNo, _DescriptionName, _DescriptionValue, _FieldColor, _RowNo });
                                    }
                                }
                                else
                                {
                                    _NewLineitem.Rows.Add(new object[] { _SrNo, _DescriptionName, _DescriptionValue, _FieldColor, _RowNo });

                                }


                            }
                        }
                    }


                    dsCopy = CopyDatatableIntoMemory(dt_Headerdata, _NewLineitem);
                }
                else
                {
                    dsCopy = CopyDatatableIntoMemory(dt_Headerdata, dt_lineItem);
                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in append NMFC value in line item--->" + ex,LogType.E);
            }
            return dsCopy;

        }



        /// <summary>
        /// update 'true' and 'false' charcter words into '1' and '0' numbers
        /// if and word occure instead of '1' or'0' then this function replace it
        /// </summary
        /// <param name="dtheader"></param>
        /// <param name="dtlineItem"></param>
        /// <returns></returns>
        public static DataSet UpdateFieldColorInDatatable(DataTable dtheader, DataTable dtlineItem)
        {
            DataSet ds_filedcolor = new DataSet();
            try
            {

                if (dtlineItem.Rows.Count > 0)
                {
                    string str_True = "True";
                    for (int i = 0; i < dtlineItem.Rows.Count; i++)
                    {
                        string FieldColorValue = Convert.ToString(dtlineItem.Rows[i]["FieldColor"]);

                        if (FieldColorValue.ToLower().Trim() == str_True.ToLower().Trim() || FieldColorValue.ToLower().Trim() == "1")
                        {
                            dtlineItem.Rows[i]["FieldColor"] = cls_AppConstant.FieldColorOne;
                        }
                        else
                        {
                            dtlineItem.Rows[i]["FieldColor"] = cls_AppConstant.FieldColorZero;
                        }

                    }
                    ds_filedcolor = CopyDatatableIntoMemory(dtheader, dtlineItem);
                    return ds_filedcolor;
                }

            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Update FieldColor value like 'True' to '1'-->" + ex,LogType.E);
                ds_filedcolor = CopyDatatableIntoMemory(dtheader, dtlineItem);
                return ds_filedcolor;
            }
            ds_filedcolor = CopyDatatableIntoMemory(dtheader, dtlineItem);
            return ds_filedcolor;
        }

        /// </summary>
        /// add speci_inst,delivary require,accssorial,quote number in lineitem from headerfield
        /// find Codename in values using hash#
        /// seprate multiple values using ~ till
        /// final soted dataset will return
        /// <param name="dtHeader"></param>
        /// <param name="dtLineItem"></param>
        /// <returns></returns>
        public static DataSet AddHeaderFieldInLineItem(DataTable dtHeader, DataTable dtLineItem)
        {
            DataSet ds = new DataSet();

            string _Description = string.Empty;
            string _CodeName = string.Empty;
            int Index = 0;
            int SrNO = 0;
            double RowNo = 0;
            try
            {
                List<string> list_temp_Header = cls_AppConstant.HeaderToLineItem();
                foreach (var item in list_temp_Header)
                {

                    DataTable dtSpecial_Instr = FindFieldRowInDataTable(dtHeader, item);
                    dtHeader = RemoveFieldRowFromDataTable(dtHeader, item);
                    if (dtSpecial_Instr.Rows.Count > 0)
                    {
                        //add if condition on top of dtSpecial_Instr for check datatble row count
                        string _FieldName = Convert.ToString(dtSpecial_Instr.Rows[0]["FieldName"]);
                        string _FieldValue = Convert.ToString(dtSpecial_Instr.Rows[0]["FieldValue"]);
                        string _FieldColor = Convert.ToString(dtSpecial_Instr.Rows[0]["FieldColor"]);
                        if (!string.IsNullOrEmpty(_FieldValue))
                        {

                            string[] TillSplitValues = TillSplitter(_FieldValue);
                            string[] TillSplitFiledColor = TillSplitter(_FieldColor);
                            if (TillSplitValues != null)
                            {
                                for (int i = 0; i < TillSplitValues.Length; i++)
                                {
                                    string[] _CodeNameValues = HashSplitterForCode(TillSplitValues[i].ToString());
                                    _Description = _CodeNameValues[0];
                                    _CodeName = _CodeNameValues[1];
                                    if (_CodeNameValues != null && _CodeNameValues.Length > 0)
                                    {
                                        if (dtLineItem.Rows.Count > 0)
                                        {
                                            Index = dtLineItem.Rows.Count - 1;
                                            SrNO = Convert.ToInt32(dtLineItem.Rows[Index]["SrNo"].ToString());
                                            RowNo = Convert.ToDouble(dtLineItem.Rows[Index]["RowNo"].ToString());
                                        }
                                        //if (_FieldName.ToLower().Trim() == cls_HeaderStaticField.Accessorial.ToLower().Trim())
                                        //{
                                        //    SrNO++;
                                        //    RowNo++;
                                        //    dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, _CodeName, TillSplitFiledColor[i].ToString(), RowNo });
                                        //    SrNO++;
                                        //    dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, "", TillSplitFiledColor[i].ToString(), RowNo });
                                        //}
                                        //else
                                        //{
                                        //List<string> _liststring = SetTextOnNewLine(_Description, cls_AppConstant.DescriptionTextLength);
                                        //if (_liststring.Count > 1)
                                        //{
                                        //    int count = 0;
                                        //    _liststring.ForEach(x =>
                                        //    {

                                        //        SrNO++;
                                        //        RowNo++;
                                        //        if (count == 0)
                                        //        {
                                        //            dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, _CodeName, TillSplitFiledColor[i].ToString(), RowNo });
                                        //            count++;
                                        //        }
                                        //        else
                                        //        {
                                        //            dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, "", TillSplitFiledColor[i].ToString(), RowNo });

                                        //        }

                                        //        SrNO++;
                                        //        dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, x, TillSplitFiledColor[i].ToString(), RowNo });
                                        //    });
                                        //}
                                        //else
                                        //{
                                        SrNO++;
                                        RowNo++;
                                        dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, _CodeName, TillSplitFiledColor[i].ToString(), RowNo });
                                        SrNO++;
                                        dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, _Description, TillSplitFiledColor[i].ToString(), RowNo });
                                        //}
                                        //}
                                    }
                                }
                            }
                            else
                            {
                                //s#s
                                string[] CodeNameValues = HashSplitterForCode(_FieldValue);
                                if (CodeNameValues != null && CodeNameValues.Length > 0)
                                {
                                    _Description = CodeNameValues[0];
                                    _CodeName = CodeNameValues[1];
                                    if (dtLineItem.Rows.Count > 0)
                                    {
                                        Index = dtLineItem.Rows.Count - 1;
                                        SrNO = Convert.ToInt32(dtLineItem.Rows[Index]["SrNo"].ToString());
                                        RowNo = Convert.ToDouble(dtLineItem.Rows[Index]["RowNo"].ToString());
                                    }

                                    //if (_FieldName.ToLower().Trim() == cls_HeaderStaticField.Accessorial.ToLower().Trim())
                                    //{
                                    //    SrNO++;
                                    //    RowNo++;
                                    //    dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, _CodeName, _FieldColor, RowNo });
                                    //    SrNO++;
                                    //    dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, "", _FieldColor, RowNo });
                                    //}
                                    //else
                                    //{
                                    //List<string> _liststring = SetTextOnNewLine(_Description, cls_AppConstant.DescriptionTextLength);
                                    //if (_liststring.Count > 1)
                                    //{
                                    //    int count = 0;
                                    //    _liststring.ForEach(x =>
                                    //    {

                                    //        SrNO++;
                                    //        RowNo++;
                                    //        if (count == 0)
                                    //        {
                                    //            dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, _CodeName, _FieldColor, RowNo });
                                    //            count++;
                                    //        }
                                    //        else
                                    //        {
                                    //            dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, "", _FieldColor, RowNo });

                                    //        }

                                    //        SrNO++;
                                    //        dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, x, _FieldColor, RowNo });
                                    //    });
                                    //}
                                    //else
                                    //{
                                    SrNO++;
                                    RowNo++;
                                    dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Class_Code, _CodeName, _FieldColor, RowNo });
                                    SrNO++;
                                    dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, _Description, _FieldColor, RowNo });
                                    //}
                                    //}
                                }
                            }
                        }
                    }
                }

                //For QUOTE number
                #region QUOTE Number
                DataTable dt_quote = FindFieldRowInDataTable(dtHeader, cls_AppConstant.Del_Quote_No);
                dtHeader = RemoveFieldRowFromDataTable(dtHeader, cls_AppConstant.Del_Quote_No);

                if (dt_quote.Rows.Count > 0)
                {
                    string Del_Quote = Convert.ToString(dt_quote.Rows[0]["FieldValue"]);
                    if (!string.IsNullOrEmpty(Del_Quote))
                    {
                        dtLineItem = AddFieldValueinLineItem(dtLineItem, cls_AppConstant.Del_Quote_Code, Del_Quote, dt_quote.Rows[0]["FieldColor"].ToString());
                    }
                }

                #endregion

                ds = CopyDatatableIntoMemory(dtHeader, dtLineItem);

                return ds;

            }
            catch (Exception ex)
            {
                //cls_Global.func_ErrorLog("Exception in AddHeaderFieldInLineItem function--->" + ex);
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in AddHeaderFieldInLineItem function--->" + ex.StackTrace,LogType.E);
                ds = CopyDatatableIntoMemory(dtHeader, dtLineItem);
                return ds;
            }

        }


        #endregion

        #region Telephone Number add Lineitems

        public static DataTable CheckHazadous(DataTable dt_input)
        {
            DataTable dt = new DataTable();
            try
            {
                var Hazardous = dt_input.AsEnumerable().Where(x => x.Field<string>("FieldValue") == cls_AppConstant.HazordousX && x.Field<string>("FieldColor") == cls_AppConstant.FieldColorOne).ToList();
                if (Hazardous.Any())
                {
                    dt = Hazardous.CopyToDataTable();
                }

            }
            catch (Exception)
            {
            }
            return dt;

        }

        public static DataSet TelephoneHazmatInLineItem(DataTable dtHeader, DataTable dtLineItem)
        {
            int Index = 0;
            int SrNO = 0;
            double RowNo = 0;
            string _ERT_Code = string.Empty;
            string _ErtFiledColor = string.Empty;
            string[] _ERTCodeArray = null;
            string[] _HashCodeArray = null;
            string[] _ErtFiledColorarray = null;
            DataSet ds = new DataSet();

            List<KeyValuePair<String, String>> _keyValuePairs = new List<KeyValuePair<String, String>>();
            try
            {

                DataTable dt_ERT = FindFieldRowInDataTable(dtHeader, cls_AppConstant.ERT_Code);
                if (dt_ERT.Rows.Count > 0)
                {
                    DataTable dt_Hazmet = FindFieldRowInDataTable(dtLineItem, LineItemStaticField.Hazmet_Code);
                    //EMER#ds~
                    //code
                    _ERT_Code = Convert.ToString(dt_ERT.Rows[0]["FieldValue"]);
                    _ErtFiledColor = Convert.ToString(dt_ERT.Rows[0]["FieldColor"]);

                    //e,c,tre
                    //1~0~1
                    //
                    if (_ERT_Code.Contains(cls_AppConstant.Till))
                    {
                        _ERTCodeArray = TillSplitter(_ERT_Code);
                        _ErtFiledColorarray = TillSplitter(_ErtFiledColor);
                        foreach (var item in _ERTCodeArray)
                        {
                            if (item.Contains(cls_AppConstant.Hash))
                            {
                                _HashCodeArray = HashSplitterForCode(item);
                                _keyValuePairs.Add(new KeyValuePair<String, String>(_HashCodeArray[1], _HashCodeArray[0]));//code,desc
                            }
                            else
                            {
                                _keyValuePairs.Add(new KeyValuePair<String, String>(item, ""));//code,desc
                            }
                        }
                    }
                    else
                    {

                        if (_ERT_Code.Contains(cls_AppConstant.Hash))
                        {
                            _HashCodeArray = HashSplitterForCode(_ERT_Code);
                            _keyValuePairs.Add(new KeyValuePair<String, String>(_HashCodeArray[1], _HashCodeArray[0]));//code,desc
                        }
                        else
                        {
                            _keyValuePairs.Add(new KeyValuePair<String, String>(_ERT_Code, ""));//code,desc
                        }
                    }
                    int count = 0;

                    if (_keyValuePairs.Count > 0)
                    {
                        _keyValuePairs.ForEach(x =>
                        {

                            if (!string.IsNullOrEmpty(_ErtFiledColor))
                            {
                                //EMER==EMER
                                //

                                if (x.Key.ToLower().Trim() == cls_AppConstant.EMER.ToLower().Trim())
                                {
                                    dt_Hazmet = CheckHazadous(dt_Hazmet);
                                    if (dt_Hazmet.Rows.Count > 0)
                                    {
                                        DataTable dt_Tel = FindFieldRowInDataTable(dtHeader, cls_AppConstant.Tel_HazMat);
                                        string _Tel_Hazmat = Convert.ToString(dt_Tel.Rows[0]["FieldValue"]);

                                        string Tel_Hazmat_color = Convert.ToString(dt_Tel.Rows[0]["FieldColor"]);
                                        if (!string.IsNullOrEmpty(Tel_Hazmat_color) && Tel_Hazmat_color.Trim() == cls_AppConstant.FieldColorOne.Trim())
                                        {
                                            dtLineItem = AddFieldValueinLineItem(dtLineItem, x.Key, cls_AppConstant.EMERDescription, cls_AppConstant.FieldColorOne);
                                            if (dtLineItem.Rows.Count > 0)
                                            {
                                                Index = dtLineItem.Rows.Count - 1;
                                                SrNO = Convert.ToInt32(dtLineItem.Rows[Index]["SrNo"].ToString());
                                                RowNo = Convert.ToDouble(dtLineItem.Rows[Index]["RowNo"].ToString());
                                            }
                                            SrNO++;
                                            RowNo++;
                                            dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, _Tel_Hazmat, cls_AppConstant.FieldColorOne, RowNo });
                                        }
                                    }
                                }
                                else
                                {
                                    if (_ERTCodeArray == null && _ErtFiledColorarray == null)
                                    {
                                        //For single FieldColorValue date 9 jan 2021
                                        if (dtLineItem.Rows.Count > 0)
                                        {
                                            Index = dtLineItem.Rows.Count - 1;
                                            SrNO = Convert.ToInt32(dtLineItem.Rows[Index]["SrNo"].ToString());
                                            RowNo = Convert.ToDouble(dtLineItem.Rows[Index]["RowNo"].ToString());
                                        }

                                        dtLineItem = AddFieldValueinLineItem(dtLineItem, x.Key, "", _ErtFiledColor);

                                    }
                                    else
                                    {
                                        //chemt~ctrec 0
                                        //for multiple Fieldcolorvalue date 9 jan 2021
                                        if (_ERTCodeArray.Length == _ErtFiledColorarray.Length)
                                        {
                                            if (dtLineItem.Rows.Count > 0)
                                            {
                                                Index = dtLineItem.Rows.Count - 1;
                                                SrNO = Convert.ToInt32(dtLineItem.Rows[Index]["SrNo"].ToString());
                                                RowNo = Convert.ToDouble(dtLineItem.Rows[Index]["RowNo"].ToString());
                                            }

                                            dtLineItem = AddFieldValueinLineItem(dtLineItem, x.Key, "", _ErtFiledColorarray[count]);
                                        }
                                    }
                                }
                            }
                            count++;
                        });
                    }
                    //}

                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in convert Telephone header to lineitems TelephoneHazmatInLineItem--->" + ex,LogType.E);
            }
            ds = CopyDatatableIntoMemory(dtHeader, dtLineItem);
            return ds;
        }


        public static DataSet TelephoneHazmatInLineItem_old(DataTable dtHeader, DataTable dtLineItem)
        {
            int Index = 0;
            int SrNO = 0;
            double RowNo = 0;
            string _ERT_Code = string.Empty;
            string _ErtFiledColor = string.Empty;
            string[] _ERTCodeArray = null;
            DataSet ds = new DataSet();
            try
            {

                DataTable dt_Temp = FindFieldRowInDataTable(dtHeader, cls_AppConstant.ERT_Code);
                if (dt_Temp.Rows.Count > 0)
                {
                    //descrip#code
                    _ERT_Code = Convert.ToString(dt_Temp.Rows[0]["FieldValue"]);
                    _ErtFiledColor = Convert.ToString(dt_Temp.Rows[0]["FieldColor"]);


                    if (!string.IsNullOrEmpty(_ERT_Code) && _ErtFiledColor == cls_AppConstant.FieldColorOne)
                    {

                        DataTable dt_Tel = FindFieldRowInDataTable(dtHeader, cls_AppConstant.Tel_HazMat);
                        if (dt_Tel.Rows.Count > 0)
                        {
                            //string[] CodeArray = TillSplitter(_ERT_Code);
                            //foreach (var item in CodeArray)
                            //{
                            //

                            _ERTCodeArray = HashSplitterForCode(_ERT_Code);//item#hdjf
                            if (_ERTCodeArray != null && _ERTCodeArray.Length > 0)
                            {
                                string _Description = _ERTCodeArray[0];
                                string _Code = _ERTCodeArray[1];

                                if (!string.IsNullOrEmpty(dt_Tel.Rows[0]["FieldValue"].ToString()))
                                {
                                    Index = dtLineItem.Rows.Count - 1;
                                    SrNO = Convert.ToInt32(dtLineItem.Rows[Index]["SrNo"].ToString());
                                    RowNo = Convert.ToDouble(dtLineItem.Rows[Index]["RowNo"].ToString());
                                }

                                dtLineItem = AddFieldValueinLineItem(dtLineItem, _Code, _Description, dt_Tel.Rows[0]["FieldColor"].ToString());

                                if (!string.IsNullOrEmpty(dt_Tel.Rows[0]["FieldValue"].ToString()))
                                {
                                    Index = dtLineItem.Rows.Count - 1;
                                    SrNO = Convert.ToInt32(dtLineItem.Rows[Index]["SrNo"].ToString());
                                    RowNo = Convert.ToDouble(dtLineItem.Rows[Index]["RowNo"].ToString());
                                }

                                SrNO++;
                                RowNo++;
                                dtLineItem.Rows.Add(new object[] { SrNO, LineItemStaticField.Description, dt_Tel.Rows[0]["FieldValue"].ToString(), dt_Tel.Rows[0]["FieldColor"], RowNo });
                                //}
                            }
                        }
                    }


                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in convert Telephone header to lineitems TelephoneHazmatInLineItem--->" + ex,LogType.E );
            }
            ds = CopyDatatableIntoMemory(dtHeader, dtLineItem);
            return ds;
        }


        public static DataSet TelephoneConsiInLineItem(DataTable dtHeader, DataTable dtLineItem)
        {

            DataTable dt_Sortedlineitem = new DataTable();
            string _Con_TelNo = string.Empty;
            DataSet ds = new DataSet();
            try
            {

                DataTable dt_Tel = FindFieldRowInDataTable(dtHeader, cls_AppConstant.Con_TelNo);
                if (dt_Tel.Rows.Count > 0)
                {
                    _Con_TelNo = Convert.ToString(dt_Tel.Rows[0]["FieldValue"]);
                    if (!string.IsNullOrEmpty(_Con_TelNo))
                    {
                        dt_Sortedlineitem = AddFieldValueinLineItem(dtLineItem, cls_AppConstant.Con_TelNo_Code, _Con_TelNo, dt_Tel.Rows[0]["FieldColor"].ToString());

                    }
                }

            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception-Exception in convert Telephone Individual header to lineitems TelephoneConsiInLineItem--->" + ex,LogType.E);
            }

            ds = CopyDatatableIntoMemory(dtHeader, dtLineItem);
            return ds;
        }
        #endregion



        #endregion

        #region Convert list to datatable and convert datatable to list 


        public static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }

        public static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                        pro.SetValue(obj, dr[column.ColumnName], null);
                    else
                        continue;
                }
            }
            return obj;
        }

        public static DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties =
                TypeDescriptor.GetProperties(typeof(T));

            DataTable table = new DataTable();

            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);

            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }
            return table;
        }
        #endregion






        #region ColorCode


        public static cls_ReturnStatus UpdateIntoFRPD010_UpdateIntoFRP001_InsertIntoFRP002(Dictionary<string, string> dicHeaderData_FRP001, DataTable dtFRP002TRU, int Linenumber, List<int> _listLinenumber_FRP002TRU, string connectionString, string ProNumber, string billingStatus,string str_DLV, OdbcParameter[] parameters = null)
        {
            cls_ReturnStatus _ReturnStatus = new cls_ReturnStatus();
            int resultVal = -1;
            // bool ISFHSTAT_PP = false;

            _ReturnStatus.Message = "Transfer Success";
            _ReturnStatus.Status = "S";
            _ReturnStatus.IsStatus = true;

            try
            {

                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == false)
                {
                    OdbcConnection db2Connection = new OdbcConnection();
                    //MessageBox.Show("GetDataTableByText_IsEDI: DB2ConnectionString: " + DB2ConnectionString);
                    db2Connection.ConnectionString = DBUtilities.DB2ConnectionString;

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
                    connection.ConnectionString = DBUtilities.DB2ConnectionString;
                    connection.Open();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("DB2Conn");
                    }
                    iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", connection);
                }

                BOLDocumentDataExtraction.DataExtraction.WriteLog("UpdateIntoFRPD010_UpdateIntoFRP001_InsertIntoFRP002()--> DB2 connection Opened", LogType.P);

                string queryString = "";

                #region"UpdateIntoFRP001"

                if (dicHeaderData_FRP001.Count > 0)
                {
                    StringBuilder sbHeaderUpdate = new StringBuilder();
                    StringBuilder sbHeader = new StringBuilder();
                    sbHeader.Append("FHPRO");
                    foreach (KeyValuePair<string, string> pairkey in dicHeaderData_FRP001)
                    {
                        sbHeader.Append("," + pairkey.Key.Trim());
                        sbHeaderUpdate.Append(pairkey.Key.Trim() + "=? ,");
                    }
                    resultVal = -1;
                    queryString = "UPDATE FRP001 SET " + Convert.ToString(sbHeaderUpdate).Trim(',').Trim() + " WHERE FHPRO =? and FHSTAT =? ";

                    if (connection != null && !string.IsNullOrEmpty(queryString))
                    {
                        using (OdbcCommand command = new OdbcCommand(queryString, connection))
                        {

                            string[] arKeyName = Convert.ToString(sbHeader).Trim().Split(',');

                            for (int ikey = 1; ikey < arKeyName.Length; ikey++)
                            {

                                var rows = dicHeaderData_FRP001.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                            == arKeyName[ikey].ToLower().Trim()).ToList();

                                if (arKeyName[ikey].Trim() == "FHTOTP" || arKeyName[ikey].Trim() == "FHSWGT")
                                {
                                    if (rows.Count > 0)
                                    {
                                        if (!string.IsNullOrEmpty(rows[0].Value))
                                        {
                                            // here you can use it safety
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = Convert.ToInt32(rows[0].Value.Trim());
                                        }
                                        else
                                        {

                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                        }
                                    }
                                    else
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                    }
                                }
                                else if (arKeyName[ikey].Trim() == "FHRADF" || arKeyName[ikey].Trim() == "FHRADT")
                                {
                                    if (rows.Count > 0)
                                    {
                                        if (!string.IsNullOrEmpty(rows[0].Value))
                                        {
                                            string Date1YYMMDD = Convert.ToString(rows[0].Value.Trim());//MMDDYY format
                                            Date1YYMMDD = Date1YYMMDD.Trim();
                                            if (Date1YYMMDD.Length == 6)
                                            {
                                                Date1YYMMDD = "1" + Date1YYMMDD.Substring(4, 2) + Date1YYMMDD.Substring(0, 2) + Date1YYMMDD.Substring(2, 2);
                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Date1YYMMDD;
                                            }
                                            else
                                            {
                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                            }
                                        }
                                        else
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                        }
                                    }
                                    else
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                    }
                                }
                                else
                                {

                                    if (rows.Count > 0)
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Convert.ToString(rows[0].Value.Trim());
                                    }
                                    else
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                    }
                                }

                            }

                            command.Parameters.Add("FHPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                            command.Parameters.Add("FHSTAT", OdbcType.VarChar).Value = billingStatus;

                            resultVal = command.ExecuteNonQuery();
                        }
                    }

                }
                else
                {
                    resultVal = 1;
                }
                #endregion


                if (resultVal <= 0)
                {
                    #region"Billing WIP In AS400"

                    _ReturnStatus.Status = "F";
                    _ReturnStatus.IsStatus = false;
                    _ReturnStatus.Message = "Billing WIP In AS400";
                    _ReturnStatus.Value = "Billing WIP In AS400 FHSTAT Status <> " + billingStatus;

                    BOLDocumentDataExtraction.DataExtraction.WriteLog("UpdateIntoFRP001 table fail.", LogType.P);
                    #endregion
                    
                }
                else
                {
                    BOLDocumentDataExtraction.DataExtraction.WriteLog("UpdateIntoFRP001 table Succ", LogType.P);

                    #region "Update OR Insert F2DLVA value FRP001EX"

                    try
                    {


                        if (!string.IsNullOrEmpty(str_DLV))
                        {
                            string queryString_FRP001EX = "select * from FRP001EX where F2PRO = " + ProNumber + " ";
                            DataTable dtFRP001EX = new DataTable();

                            if (connection != null && !string.IsNullOrEmpty(queryString_FRP001EX))
                            {
                                using (OdbcCommand cmd = new OdbcCommand(queryString_FRP001EX, connection))
                                {
                                    cmd.CommandType = CommandType.Text;
                                    if (parameters != null)
                                    {
                                        cmd.Parameters.AddRange(parameters);
                                    }
                                    using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                                    {
                                        da.Fill(dtFRP001EX);

                                    }
                                }
                            }


                            if (dtFRP001EX.Rows.Count > 0)
                            {
                                queryString = "UPDATE FRP001EX SET F2DLVA =?  WHERE F2PRO =?";
                            }
                            else
                            {
                                queryString = "INSERT INTO FRP001EX (F2DLVA, F2PRO) VALUES (?,?)";
                            }

                            if (connection != null && !string.IsNullOrEmpty(queryString))
                            {
                                using (OdbcCommand command = new OdbcCommand(queryString, connection))
                                {
                                    command.Parameters.Add("F2DLVA", OdbcType.VarChar).Value = str_DLV;
                                    command.Parameters.Add("F2PRO", OdbcType.VarChar).Value = ProNumber;

                                    command.ExecuteNonQuery(); ;
                                }
                            }

                        }
                    }
                    catch (Exception EFRP001EX)
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("FRP001EX table fail--->" + EFRP001EX, LogType.P);
                    }
                    #endregion



                    #region"INSERT FRP002TRU"

                    if (dtFRP002TRU != null && dtFRP002TRU.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtFRP002TRU.Rows)
                        {
                            Linenumber = Linenumber + 1;
                            resultVal = -1;
                            try
                            {

                                queryString = "INSERT INTO FRP002TRU (FDPRO,FDLIN,FDPCS,FDCMCL,FDHAZ,FDPKGC,FDDES,FDWGT,FDUID) VALUES (?,?,?,?,?,?,?,?,?)";

                                if (connection != null && !string.IsNullOrEmpty(queryString))
                                {
                                    using (OdbcCommand command = new OdbcCommand(queryString, connection))
                                    {
                                        string Pices = Convert.ToString(dr["PCS"]).Trim();
                                        string Weight = Convert.ToString(dr["Weight"]).Trim();

                                        command.Parameters.Add("FDPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                                        command.Parameters.Add("FDLIN", OdbcType.Int).Value = Linenumber;

                                        if (!string.IsNullOrEmpty(Pices))
                                        {
                                            // here you can use it safety
                                            command.Parameters.Add("FDPCS", OdbcType.Int).Value = Convert.ToInt32(Pices);
                                        }
                                        else
                                        {

                                            command.Parameters.Add("FDPCS", OdbcType.Int).Value = 0;//DBNull.Value;
                                        }
                                        command.Parameters.Add("FDCMCL", OdbcType.VarChar).Value = Convert.ToString(dr["Code"]);
                                        command.Parameters.Add("FDHAZ", OdbcType.VarChar).Value = Convert.ToString(dr["H"]);
                                        command.Parameters.Add("FDPKGC", OdbcType.VarChar).Value = Convert.ToString(dr["P"]);
                                        command.Parameters.Add("FDDES", OdbcType.VarChar).Value = Convert.ToString(dr["Description"]);

                                        if (!string.IsNullOrEmpty(Weight))
                                        {
                                            command.Parameters.Add("FDWGT", OdbcType.Int).Value = Convert.ToInt32(Weight);
                                        }
                                        else
                                        {
                                            command.Parameters.Add("FDWGT", OdbcType.Int).Value = 0;//DBNull.Value;
                                        }

                                        command.Parameters.Add("FDUID", OdbcType.VarChar).Value = "TRUBOT";

                                        resultVal = command.ExecuteNonQuery();

                                    }
                                }

                            }
                            catch (Exception ex002)
                            {
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("InsertIntoFRP002TRU table fail--->" + ex002, LogType.P);
                            }

                            if (resultVal <= 0)
                            {
                                try
                                {

                                    foreach (DataRow drRow in dtSortedTable.Select("LineNumber = '" + Convert.ToString(Linenumber) + "'"))
                                        drRow.Delete();
                                }
                                catch (Exception e)
                                {
                                    BOLDocumentDataExtraction.DataExtraction.WriteLog("InsertIntoFRP002TRU: Delete Row from FRPD010--> for Row " + Linenumber + " Error: " + e, LogType.E);
                                }

                            }
                        }
                    }
                    else
                    {
                        resultVal = 1;
                    }
                    #endregion

                    if (resultVal <= 0)
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("IntoFRP002TRU table fail for Row :" + Linenumber, LogType.P);
                    }
                    else
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("IntoFRP002TRU table fail Succ", LogType.P);
                    }

                    #region"INSERT INTO FRPD010"
                    if (dtSortedTable != null && dtSortedTable.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtSortedTable.Rows.Count; i++)
                        {
                            queryString = "INSERT INTO FRPD010 (HXDPRO, HXDFIL, HXDFLD, HXDLIN) VALUES (?,?,?,?)";
                            if (connection != null && !string.IsNullOrEmpty(queryString))
                            {
                                using (OdbcCommand command = new OdbcCommand(queryString, connection))
                                {
                                    command.Parameters.Add("HXDPRO", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["ProNumber"];
                                    command.Parameters.Add("HXDFIL", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["FileName"];
                                    command.Parameters.Add("HXDFLD", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["FieldName"];
                                    command.Parameters.Add("HXDLIN", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["LineNumber"];
                                    resultVal = command.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                    #endregion

                }

                if (connection.State == ConnectionState.Open)
                { connection.Close(); }

            }
            catch (Exception ex)
            {
                _ReturnStatus.Status = "F";
                _ReturnStatus.IsStatus = false;
                _ReturnStatus.Message = "Transfer Failed";
                _ReturnStatus.Value = ex.Message;
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in UpdateIntoFRPD010_UpdateIntoFRP001_InsertIntoFRP002()--->" + ex, LogType.E);
            }


            return _ReturnStatus;

        }
        public static cls_ReturnStatus UpdateIntoFRPD010_UpdateIntoFRP001_InsertIntoFRP002_Backup_14Jun21(Dictionary<string, string> dicHeaderData_FRP001, DataTable dtFRP002TRU, int Linenumber, List<int> _listLinenumber_FRP002TRU, string connectionString, string ProNumber, string billingStatus, OdbcParameter[] parameters = null)
        {
            cls_ReturnStatus _ReturnStatus = new cls_ReturnStatus();
            int resultVal = -1;
            bool ISFHSTAT_PP = false;

            _ReturnStatus.Message = "Transfer Success";
            _ReturnStatus.Status = "S";
            _ReturnStatus.IsStatus = true;

            try
            {
                //cls_FieldRuleValidation.dtSortedTable
               

                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == false)
                {
                    OdbcConnection db2Connection = new OdbcConnection();
                    //MessageBox.Show("GetDataTableByText_IsEDI: DB2ConnectionString: " + DB2ConnectionString);
                    db2Connection.ConnectionString = DBUtilities.DB2ConnectionString;

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
                    connection.ConnectionString = DBUtilities.DB2ConnectionString;
                    connection.Open();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("DB2Conn");
                    }
                    iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", connection);
                }



                #region"CHECK BillingStatus FHSTAT = 'P'"
                string queryString = "select FHPRO, FHSTAT from FRP001 where FHPRO = " + ProNumber + " and FHSTAT = '" + billingStatus + "' ";
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

                    
                    #region"UpdateIntoFRP001"

                    StringBuilder sbHeaderUpdate = new StringBuilder();
                    StringBuilder sbHeader = new StringBuilder();
                    sbHeader.Append("FHPRO");
                    foreach (KeyValuePair<string, string> pairkey in dicHeaderData_FRP001)
                    {
                        sbHeader.Append("," + pairkey.Key.Trim());
                        sbHeaderUpdate.Append(pairkey.Key.Trim() + "=? ,");
                    }
                    resultVal = -1;
                    queryString = "UPDATE FRP001 SET " + Convert.ToString(sbHeaderUpdate).Trim(',').Trim() + " WHERE FHPRO =? and FHSTAT =? ";

                    if (connection != null && !string.IsNullOrEmpty(queryString))
                    {
                        using (OdbcCommand command = new OdbcCommand(queryString, connection))
                        {

                            string[] arKeyName = Convert.ToString(sbHeader).Trim().Split(',');

                            for (int ikey = 1; ikey < arKeyName.Length; ikey++)
                            {

                                var rows = dicHeaderData_FRP001.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                            == arKeyName[ikey].ToLower().Trim()).ToList();

                                if (arKeyName[ikey].Trim() == "FHTOTP" || arKeyName[ikey].Trim() == "FHSWGT")
                                {
                                    if (rows.Count > 0)
                                    {
                                        if (!string.IsNullOrEmpty(rows[0].Value))
                                        {
                                            // here you can use it safety
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = Convert.ToInt32(rows[0].Value.Trim());
                                        }
                                        else
                                        {

                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                        }
                                    }
                                    else
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                    }
                                }
                                else if (arKeyName[ikey].Trim() == "FHRADF" || arKeyName[ikey].Trim() == "FHRADT")
                                {
                                    if (rows.Count > 0)
                                    {
                                        if (!string.IsNullOrEmpty(rows[0].Value))
                                        {
                                            string Date1YYMMDD = Convert.ToString(rows[0].Value.Trim());//MMDDYY format
                                            Date1YYMMDD = Date1YYMMDD.Trim();
                                            if (Date1YYMMDD.Length == 6)
                                            {
                                                Date1YYMMDD = "1" + Date1YYMMDD.Substring(4, 2) + Date1YYMMDD.Substring(0, 2) + Date1YYMMDD.Substring(2, 2);
                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Date1YYMMDD;
                                            }
                                            else
                                            {
                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                            }
                                        }
                                        else
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                        }
                                    }
                                    else
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                    }
                                }
                                else
                                {

                                    if (rows.Count > 0)
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Convert.ToString(rows[0].Value.Trim());
                                    }
                                    else
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                    }
                                }

                            }

                            command.Parameters.Add("FHPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                            command.Parameters.Add("FHSTAT", OdbcType.VarChar).Value = billingStatus;

                            resultVal = command.ExecuteNonQuery();
                        }
                    }

                    #endregion


                    if (resultVal <= 0)
                    {
                        _ReturnStatus.Message = "Transfer Failed";
                        _ReturnStatus.Status = "F";
                        _ReturnStatus.IsStatus = false;
                        _ReturnStatus.Value = "UpdateIntoFRP001 table fail";
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("UpdateIntoFRP001 table fail!!", LogType.P);
                    }
                    else
                    {
                        BOLDocumentDataExtraction.DataExtraction.WriteLog("UpdateIntoFRP001 table Succ", LogType.P);
                        #region"INSERT FRP002TRU"

                        if (dtFRP002TRU != null && dtFRP002TRU.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dtFRP002TRU.Rows)
                            {
                                Linenumber = Linenumber + 1;
                                resultVal = -1;
                                try
                                {

                                    queryString = "INSERT INTO FRP002TRU (FDPRO,FDLIN,FDPCS,FDCMCL,FDHAZ,FDPKGC,FDDES,FDWGT,FDUID) VALUES (?,?,?,?,?,?,?,?,?)";

                                    if (connection != null && !string.IsNullOrEmpty(queryString))
                                    {
                                        using (OdbcCommand command = new OdbcCommand(queryString, connection))
                                        {
                                            string Pices = Convert.ToString(dr["PCS"]).Trim();
                                            string Weight = Convert.ToString(dr["Weight"]).Trim();

                                            command.Parameters.Add("FDPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                                            command.Parameters.Add("FDLIN", OdbcType.Int).Value = Linenumber;

                                            if (!string.IsNullOrEmpty(Pices))
                                            {
                                                // here you can use it safety
                                                command.Parameters.Add("FDPCS", OdbcType.Int).Value = Convert.ToInt32(Pices);
                                            }
                                            else
                                            {

                                                command.Parameters.Add("FDPCS", OdbcType.Int).Value = 0;//DBNull.Value;
                                            }
                                            command.Parameters.Add("FDCMCL", OdbcType.VarChar).Value = Convert.ToString(dr["Code"]);
                                            command.Parameters.Add("FDHAZ", OdbcType.VarChar).Value = Convert.ToString(dr["H"]);
                                            command.Parameters.Add("FDPKGC", OdbcType.VarChar).Value = Convert.ToString(dr["P"]);
                                            command.Parameters.Add("FDDES", OdbcType.VarChar).Value = Convert.ToString(dr["Description"]);

                                            if (!string.IsNullOrEmpty(Weight))
                                            {
                                                command.Parameters.Add("FDWGT", OdbcType.Int).Value = Convert.ToInt32(Weight);
                                            }
                                            else
                                            {
                                                command.Parameters.Add("FDWGT", OdbcType.Int).Value = 0;//DBNull.Value;
                                            }

                                            command.Parameters.Add("FDUID", OdbcType.VarChar).Value = "TRUBOT";

                                            resultVal = command.ExecuteNonQuery();

                                        }
                                    }

                                }
                                catch (Exception ex002)
                                {
                                    BOLDocumentDataExtraction.DataExtraction.WriteLog("InsertIntoFRP002TRU table fail--->" + ex002, LogType.P);
                                }

                                if (resultVal <= 0)
                                {
                                    try
                                    {

                                        foreach (DataRow drRow in dtSortedTable.Select("LineNumber = '" + Convert.ToString(Linenumber) + "'"))
                                            drRow.Delete();
                                    }
                                    catch (Exception e)
                                    {
                                        BOLDocumentDataExtraction.DataExtraction.WriteLog("Delete Row from FRPD010--> for Row " + Linenumber + " Error: " + e, LogType.E);
                                    }

                                }
                            }
                        }
                        #endregion

                        if (resultVal <= 0)
                        {
                            BOLDocumentDataExtraction.DataExtraction.WriteLog("Insert IntoFRP002TRU table fail for Row :" + Linenumber, LogType.P);
                        }

                        #region"INSERT INTO FRPD010"
                        if (dtSortedTable != null && dtSortedTable.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtSortedTable.Rows.Count; i++)
                            {
                                queryString = "INSERT INTO FRPD010 (HXDPRO, HXDFIL, HXDFLD, HXDLIN) VALUES (?,?,?,?)";
                                if (connection != null && !string.IsNullOrEmpty(queryString))
                                {
                                    using (OdbcCommand command = new OdbcCommand(queryString, connection))
                                    {
                                        command.Parameters.Add("HXDPRO", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["ProNumber"];
                                        command.Parameters.Add("HXDFIL", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["FileName"];
                                        command.Parameters.Add("HXDFLD", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["FieldName"];
                                        command.Parameters.Add("HXDLIN", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["LineNumber"];
                                        resultVal = command.ExecuteNonQuery();
                                    }
                                }
                            }
                        }
                        #endregion

                    }
                }
                else
                {

                    #region"RETURN F AS 'Billing WIP In AS400'"

                    _ReturnStatus.Status = "F";
                    _ReturnStatus.IsStatus = false;
                    _ReturnStatus.Message = "Billing WIP In AS400";
                    _ReturnStatus.Value = "Billing WIP In AS400 FHSTAT Status <> " + billingStatus;

                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Currently Billing WIP In AS400", LogType.P);
                    // WriteLog(documentName + ":" + " Billed in AS400", LogType.E);
                    //InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " Billed in AS400", executionServiceName, "PerformDataExtraction_KM2");
                    #endregion
                }


                if (connection.State == ConnectionState.Open)
                { connection.Close(); }

            }
            catch (Exception ex)
            {
                _ReturnStatus.Status = "F";
                _ReturnStatus.IsStatus = false;
                _ReturnStatus.Message = "Transfer Failed";
                _ReturnStatus.Value = ex.Message;
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in UpdateIntoFRPD010_UpdateIntoFRP001_InsertIntoFRP002()--->" + ex, LogType.E);
            }
            

            return _ReturnStatus;

        }

        public static cls_ReturnStatus UpdateIntoFRPD010_UpdateIntoFRP001_InsertIntoFRP002_BackUp19MAY21(Dictionary<string, string> dicHeaderData_FRP001, DataTable dtFRP002TRU, int Linenumber, List<int> _listLinenumber_FRP002TRU, string connectionString, string ProNumber, string billingStatus, OdbcParameter[] parameters = null)
        {
            cls_ReturnStatus _ReturnStatus = new cls_ReturnStatus();
            int resultVal = -1;
            bool ISFHSTAT_PP = false;

            _ReturnStatus.Message = "Transfer Success";
            _ReturnStatus.Status = "S";
            _ReturnStatus.IsStatus = true;

            try
            {
                //cls_FieldRuleValidation.dtSortedTable
                string queryString = "select FHPRO, FHSTAT from FRP001 where FHPRO = " + ProNumber + " and FHSTAT = '" + billingStatus + "' ";

                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == false)
                {
                    OdbcConnection db2Connection = new OdbcConnection();
                    //MessageBox.Show("GetDataTableByText_IsEDI: DB2ConnectionString: " + DB2ConnectionString);
                    db2Connection.ConnectionString = DBUtilities.DB2ConnectionString;

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
                    connection.ConnectionString = DBUtilities.DB2ConnectionString;
                    connection.Open();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("DB2Conn");
                    }
                    iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", connection);
                }



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
                    #region"INSERT INTO FRPD010"
                    if (dtSortedTable != null && dtSortedTable.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtSortedTable.Rows.Count; i++)
                        {
                            queryString = "INSERT INTO FRPD010 (HXDPRO, HXDFIL, HXDFLD, HXDLIN) VALUES (?,?,?,?)";
                            if (connection != null && !string.IsNullOrEmpty(queryString))
                            {
                                using (OdbcCommand command = new OdbcCommand(queryString, connection))
                                {
                                    command.Parameters.Add("HXDPRO", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["ProNumber"];
                                    command.Parameters.Add("HXDFIL", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["FileName"];
                                    command.Parameters.Add("HXDFLD", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["FieldName"];
                                    command.Parameters.Add("HXDLIN", OdbcType.VarChar).Value = dtSortedTable.Rows[i]["LineNumber"];
                                    resultVal = command.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                    #endregion

                    if (resultVal <= 0)
                    {
                        _ReturnStatus.Message = "Transfer Failed";
                        _ReturnStatus.Status = "F";
                        _ReturnStatus.IsStatus = false;

                        BOLDocumentDataExtraction.DataExtraction.WriteLog("InsertIntoFRPD010 table fail!!", LogType.P);
                    }
                    else
                    {
                        #region"UpdateIntoFRP001"

                        StringBuilder sbHeaderUpdate = new StringBuilder();
                        StringBuilder sbHeader = new StringBuilder();
                        sbHeader.Append("FHPRO");
                        foreach (KeyValuePair<string, string> pairkey in dicHeaderData_FRP001)
                        {
                            sbHeader.Append("," + pairkey.Key.Trim());
                            sbHeaderUpdate.Append(pairkey.Key.Trim() + "=? ,");
                        }

                        queryString = "UPDATE FRP001 SET " + Convert.ToString(sbHeaderUpdate).Trim(',').Trim() + " WHERE FHPRO =? and FHSTAT =? ";

                        if (connection != null && !string.IsNullOrEmpty(queryString))
                        {
                            using (OdbcCommand command = new OdbcCommand(queryString, connection))
                            {

                                string[] arKeyName = Convert.ToString(sbHeader).Trim().Split(',');

                                for (int ikey = 1; ikey < arKeyName.Length; ikey++)
                                {

                                    var rows = dicHeaderData_FRP001.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                                == arKeyName[ikey].ToLower().Trim()).ToList();

                                    if (arKeyName[ikey].Trim() == "FHTOTP" || arKeyName[ikey].Trim() == "FHSWGT")
                                    {
                                        if (rows.Count > 0)
                                        {
                                            if (!string.IsNullOrEmpty(rows[0].Value))
                                            {
                                                // here you can use it safety
                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = Convert.ToInt32(rows[0].Value.Trim());
                                            }
                                            else
                                            {

                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                            }
                                        }
                                        else
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                        }
                                    }
                                    else if (arKeyName[ikey].Trim() == "FHRADF" || arKeyName[ikey].Trim() == "FHRADT")
                                    {
                                        if (rows.Count > 0)
                                        {
                                            if (!string.IsNullOrEmpty(rows[0].Value))
                                            {
                                                string Date1YYMMDD = Convert.ToString(rows[0].Value.Trim());//MMDDYY format
                                                Date1YYMMDD = Date1YYMMDD.Trim();
                                                if (Date1YYMMDD.Length == 6)
                                                {
                                                    Date1YYMMDD = "1" + Date1YYMMDD.Substring(4, 2) + Date1YYMMDD.Substring(0, 2) + Date1YYMMDD.Substring(2, 2);
                                                    command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Date1YYMMDD;
                                                }
                                                else
                                                {
                                                    command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                                }
                                            }
                                            else
                                            {
                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                            }
                                        }
                                        else
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                        }
                                    }
                                    else
                                    {

                                        if (rows.Count > 0)
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Convert.ToString(rows[0].Value.Trim());
                                        }
                                        else
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                        }
                                    }

                                }

                                command.Parameters.Add("FHPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                                command.Parameters.Add("FHSTAT", OdbcType.VarChar).Value = billingStatus;

                                resultVal = command.ExecuteNonQuery();
                            }
                        }

                        #endregion

                        if (resultVal <= 0)
                        {
                            _ReturnStatus.Message = "Transfer Failed";
                            _ReturnStatus.Status = "F";
                            _ReturnStatus.IsStatus = false;

                            BOLDocumentDataExtraction.DataExtraction.WriteLog("UpdateIntoFRP001 table fail!!", LogType.P);
                        }
                        else
                        {
                            #region"INSERT FRP002TRU"
                            if (dtFRP002TRU != null && dtFRP002TRU.Rows.Count > 0)
                            {
                                foreach (DataRow dr in dtFRP002TRU.Rows)
                                {
                                    Linenumber = Linenumber + 1;
                                    queryString = "INSERT INTO FRP002TRU (FDPRO,FDLIN,FDPCS,FDCMCL,FDHAZ,FDPKGC,FDDES,FDWGT) VALUES (?,?,?,?,?,?,?,?)";
                                    if (connection != null && !string.IsNullOrEmpty(queryString))
                                    {
                                        using (OdbcCommand command = new OdbcCommand(queryString, connection))
                                        {
                                            string Pices = Convert.ToString(dr["PCS"]).Trim();
                                            string Weight = Convert.ToString(dr["Weight"]).Trim();

                                            command.Parameters.Add("FDPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                                            command.Parameters.Add("FDLIN", OdbcType.Int).Value = Linenumber;

                                            if (!string.IsNullOrEmpty(Pices))
                                            {
                                                // here you can use it safety
                                                command.Parameters.Add("FDPCS", OdbcType.Int).Value = Convert.ToInt32(Pices);
                                            }
                                            else
                                            {

                                                command.Parameters.Add("FDPCS", OdbcType.Int).Value = 0;//DBNull.Value;
                                            }
                                            command.Parameters.Add("FDCMCL", OdbcType.VarChar).Value = Convert.ToString(dr["Code"]);
                                            command.Parameters.Add("FDHAZ", OdbcType.VarChar).Value = Convert.ToString(dr["H"]);
                                            command.Parameters.Add("FDPKGC", OdbcType.VarChar).Value = Convert.ToString(dr["P"]);
                                            command.Parameters.Add("FDDES", OdbcType.VarChar).Value = Convert.ToString(dr["Description"]);
                                            if (!string.IsNullOrEmpty(Weight))
                                            {
                                                command.Parameters.Add("FDWGT", OdbcType.Int).Value = Convert.ToInt32(Weight);
                                            }
                                            else
                                            {
                                                command.Parameters.Add("FDWGT", OdbcType.Int).Value = 0;//DBNull.Value;
                                            }
                                            resultVal = command.ExecuteNonQuery();

                                        }
                                    }
                                }
                            }
                            #endregion

                            if (resultVal <= 0)
                            {
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("InsertIntoFRP002TRU table fail!!", LogType.P);
                            }
                            else
                            {
                                #region"UPDATE TRUBOT INTO FRP002TRU"
                                foreach (var _itemlinenumber in _listLinenumber_FRP002TRU)
                                {
                                    queryString = "UPDATE FRP002TRU SET FDUID =? WHERE FDPRO =?  AND FDLIN =?";

                                    if (connection != null && !string.IsNullOrEmpty(queryString))
                                    {
                                        using (OdbcCommand command = new OdbcCommand(queryString, connection))
                                        {
                                            command.CommandType = CommandType.Text;
                                            command.Parameters.Add("FDUID", OdbcType.VarChar).Value = "TRUBOT";
                                            command.Parameters.Add("FDPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                                            command.Parameters.Add("FDLIN", OdbcType.Int).Value = _itemlinenumber;

                                            resultVal = command.ExecuteNonQuery();

                                        }
                                    }
                                }
                                #endregion
                            }
                        }

                    }
                }
                else
                {

                    #region"RETURN F AS 'Billing WIP In AS400'"

                    _ReturnStatus.Status = "F";
                    _ReturnStatus.IsStatus = false;
                    _ReturnStatus.Message = "Billing WIP In AS400";

                    BOLDocumentDataExtraction.DataExtraction.WriteLog("Currently Billing WIP In AS400", LogType.P);
                    // WriteLog(documentName + ":" + " Billed in AS400", LogType.E);
                    //InsertErrorLog(documentID, ErrorType.S, documentName + ":" + " Billed in AS400", executionServiceName, "PerformDataExtraction_KM2");
                    #endregion
                }


                if (connection.State == ConnectionState.Open)
                { connection.Close(); }

            }
            catch (Exception ex)
            {
                _ReturnStatus.Status = "F";
                _ReturnStatus.IsStatus = false;
                _ReturnStatus.Message = "Transfer Failed";
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in UpdateIntoFRPD010_UpdateIntoFRP001_InsertIntoFRP002()--->" + ex, LogType.E);
            }

            return _ReturnStatus;

        }

        /// <summary>
        /// Method to insert color code data into AS400 FRPD010 table
        /// </summary>
        /// <param name="dtHeaderFldDtls"></param>
        /// <param name="ProNumber"></param>
        /// <param name="ConnectionString"></param>
        /// <returns></returns>
        public static bool InsertHeaderColorCodeIntoFRPD010_NotUse(string _FieldName, string ProNumber, string AS400DbConnectionString)
        {
            //int resultVal = 0;

            //Dictionary<string, string> dicHdFldMapping = _HeaderFieldsMappingToAS400();
            if (cls_AppConstant.dicHdFldMapping == null || cls_AppConstant.dicHdFldMapping.Count <= 0)
            {
                cls_AppConstant.dicHdFldMapping = cls_AppConstant._HeaderFieldsMappingToAS400();
            }

            //System.Text.StringBuilder sbHeaderFieldXml = new System.Text.StringBuilder();
            //DataTable dtSortedTable = new DataTable();
            //dtSortedTable.Columns.Add("ProNumber");
            //dtSortedTable.Columns.Add("FileName");
            //dtSortedTable.Columns.Add("FieldName");
            //dtSortedTable.Columns.Add("LineNumber");
            try
            {
                if (!string.IsNullOrEmpty(_FieldName))
                {


                    var rows = cls_AppConstant.dicHdFldMapping.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                            == _FieldName.ToLower().Trim()).ToList();
                    if (rows.Count > 0)
                    {

                        dtSortedTable.Rows.Add(new string[] { ProNumber.ToString(), "FRP001", rows[0].Value.ToString(), "0" });
                    }

                    return true;

                    //if (dtSortedTable != null && dtSortedTable.Rows.Count > 0)
                    //{
                    //    resultVal = InsertIntoFRPD010(dtSortedTable, AS400DbConnectionString);
                    //    if (resultVal <= 0)
                    //    { return false; }
                    //    return false;
                    //}

                }

            }
            catch (Exception ex)
            {

                return false;
            }
            finally

            {
                //dtSortedTable.Dispose();
            }
            return true;
        }


        public static string HeaderItemMappingValue(string _FieldName, string ProNumber)
        {
            string HeaderMappingValue = string.Empty;
            
            try
            {
                if (cls_AppConstant.dicHdFldMapping == null || cls_AppConstant.dicHdFldMapping.Count <= 0)
                {
                    cls_AppConstant.dicHdFldMapping = cls_AppConstant._HeaderFieldsMappingToAS400();
                }

                if (!string.IsNullOrEmpty(_FieldName))
                {
                    var rows = cls_AppConstant.dicHdFldMapping.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                            == _FieldName.ToLower().Trim()).ToList();
                    if (rows.Count > 0)
                    {
                        HeaderMappingValue = rows[0].Value.ToString();
                        dtSortedTable.Rows.Add(new string[] { ProNumber.ToString(), "FRP001", HeaderMappingValue, "0" });
                    }

                }

            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in HeaderItemMappingValue()--->" + ex, LogType.E);
                //return false;
            }

            return HeaderMappingValue;
        }

        /// <summary>
        /// Method to insert color code data into AS400 FRPD010 table
        /// </summary>
        /// <param name="dtLineFldDtls"></param>
        /// <param name="ProNumber"></param>
        /// <param name="ConnectionString"></param>
        /// <returns></returns>
        public static bool InsertLineColorCodeIntoFRPD010(string _FieldName, string ProNumber, string AS400DbConnectionString, int RowNo)
        {
            //int resultVal = 0;

            //Dictionary<string, string> dicHdFldMapping = _HeaderFieldsMappingToAS400();
            if (cls_AppConstant.dicLineFieldsMapping == null || cls_AppConstant.dicLineFieldsMapping.Count <= 0)
            {
                cls_AppConstant.dicLineFieldsMapping = cls_AppConstant._LineFieldsMappingToAS400();
            }
            //System.Text.StringBuilder sbHeaderFieldXml = new System.Text.StringBuilder();

            //DataTable dtSortedTable = new DataTable();
            //dtSortedTable.Columns.Add("ProNumber");
            //dtSortedTable.Columns.Add("FileName");
            //dtSortedTable.Columns.Add("FieldName");
            //dtSortedTable.Columns.Add("LineNumber");
            
            try
            {
                if (!string.IsNullOrEmpty(_FieldName))
                {
                    var rows = cls_AppConstant.dicLineFieldsMapping.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                            == _FieldName.ToString().ToLower().Trim()).ToList();
                    if (rows.Count > 0)
                    {

                        dtSortedTable.Rows.Add(new string[] { ProNumber.ToString(), "FRP002", rows[0].Value.ToString(), Convert.ToString(RowNo) });
                    }

                    return true;
                    //if (dtSortedTable != null && dtSortedTable.Rows.Count > 0)
                    //{
                    //    resultVal = InsertIntoFRPD010(dtSortedTable, AS400DbConnectionString);
                    //    if (resultVal <= 0)
                    //    { return false; }
                    //}
                }
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in InsertLineColorCodeIntoFRPD010()--->" + ex, LogType.E);
                return false;
            }
            finally
            {
                //dtSortedTable.Dispose();

            }
            return true;
        }

        public static int InsertIntoFRPD010(DataTable dtFRPD010, string connectionString)
        {
            int resultVal = 0;
            string queryString = "";
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    if (dtFRPD010 != null && dtFRPD010.Rows.Count > 0)
                    {
                        if (connection.State == ConnectionState.Closed)
                        { connection.Open(); }


                        for (int i = 0; i < dtFRPD010.Rows.Count; i++)
                        {
                            queryString = "INSERT INTO FRPD010 (HXDPRO, HXDFIL, HXDFLD, HXDLIN) VALUES (?,?,?,?)";
                            if (connection != null && !string.IsNullOrEmpty(queryString))
                            {
                                using (OdbcCommand command = new OdbcCommand(queryString, connection))
                                {
                                    command.Parameters.Add("HXDPRO", OdbcType.VarChar).Value = dtFRPD010.Rows[i]["ProNumber"];
                                    command.Parameters.Add("HXDFIL", OdbcType.VarChar).Value = dtFRPD010.Rows[i]["FileName"];
                                    command.Parameters.Add("HXDFLD", OdbcType.VarChar).Value = dtFRPD010.Rows[i]["FieldName"];
                                    command.Parameters.Add("HXDLIN", OdbcType.VarChar).Value = dtFRPD010.Rows[i]["LineNumber"];
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
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in InsertIntoFRPD010()--->" + ex, LogType.E);
                return resultVal;
            }
        }

        public static int InsertIntoFRP001(Dictionary<string, string> dicHeaderData, string connectionString, string ProNumber)
        {
            int resultVal = 0;
            string queryString = "";
            string HeaderValue = "";
            //string strKeyName = "FHPRO,FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC";
            //string strKeyName = "FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC";

            string strKeyName = "FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHTRM,FHSWGT,FHRADF,FHSNUM,FHDC";
            try
            {
                StringBuilder sbHeader = new StringBuilder();
                StringBuilder sbHeaderValue = new StringBuilder();
                sbHeader.Append("FHPRO");
                sbHeaderValue.Append("?");
                foreach (KeyValuePair<string, string> pairkey in dicHeaderData)
                {
                    sbHeader.Append("," + pairkey.Key.Trim());
                    sbHeaderValue.Append(",?");
                }

                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    //if (dr != null)
                    //{

                    if (connection.State == ConnectionState.Closed)
                    {
                        connection.Open();
                    }
                    //queryString = "INSERT INTO FRP001 (" + Convert.ToString(sbHeader).Trim() + ") VALUES (" + Convert.ToString(sbHeaderValue).Trim() + ")";
                     queryString = "INSERT INTO FRP001 (FHPRO,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHTRM,FHSWGT,FHRADF,FHSNUM,FHDC) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)";
                    if (connection != null && !string.IsNullOrEmpty(queryString))
                    {
                        using (OdbcCommand command = new OdbcCommand(queryString, connection))
                        {
                            command.Parameters.Add("FHPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                            string[] arKeyName = strKeyName.Trim().Split(',');



                           // foreach (KeyValuePair<string, string> pair in dicHeaderData)
                            //foreach (KeyValuePair<string, string> pair in dicHeaderData)
                            for (int ikey = 0; ikey < arKeyName.Length ; ikey++)
                            {

                                var rows = dicHeaderData.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                            == arKeyName[ikey].ToLower().Trim()).ToList();

                                if (arKeyName[ikey].Trim() == "FHTOTP" || arKeyName[ikey].Trim() == "FHSWGT")
                                {
                                    if (rows.Count > 0)
                                    {
                                        if (!string.IsNullOrEmpty(rows[0].Value))
                                        {
                                            // here you can use it safety
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = Convert.ToInt32(rows[0].Value.Trim());
                                        }
                                        else
                                        {

                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                        }
                                    }
                                    else
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                    }
                                }
                                else
                                {

                                    if (rows.Count > 0)
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Convert.ToString(rows[0].Value.Trim());
                                    }
                                    else
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                    }
                                }

                            }

                            resultVal = command.ExecuteNonQuery();
                        }
                    }

                    //}
                    if (connection.State == ConnectionState.Open)
                    { connection.Close(); }
                }
                return resultVal;
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in InsertIntoFRP001()--->" + ex, LogType.E);
                return resultVal;
            }
        }

        public static int InsertIntoFRP001TEST(Dictionary<string, string> dicHeaderData, string connectionString, string ProNumber)
        {
            int resultVal = 0;
            string queryString = "";
            string HeaderValue = "";
            //string strKeyName = "FHPRO,FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC";
            //string strKeyName = "FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC";
            try
            {
                StringBuilder sbHeader = new StringBuilder();
                StringBuilder sbHeaderValue = new StringBuilder();
                sbHeader.Append("FHPRO");
                sbHeaderValue.Append("?");
                foreach (KeyValuePair<string, string> pairkey in dicHeaderData)
                {
                    sbHeader.Append("," + pairkey.Key.Trim());
                    sbHeaderValue.Append(",?");
                }

                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    //if (dr != null)
                    //{

                    if (connection.State == ConnectionState.Closed)
                    {
                        connection.Open();
                    }
                    queryString = "INSERT INTO FRP001 (" + Convert.ToString(sbHeader).Trim() + ") VALUES (" + Convert.ToString(sbHeaderValue).Trim() + ")";
                    // queryString = "INSERT INTO FRP001 (FHPRO,FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
                    if (connection != null && !string.IsNullOrEmpty(queryString))
                    {
                        using (OdbcCommand command = new OdbcCommand(queryString, connection))
                        {
                            command.Parameters.Add("FHPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                            string[] arKeyName = Convert.ToString(sbHeader).Trim().Split(',');

                            // foreach (KeyValuePair<string, string> pair in dicHeaderData)
                            //foreach (KeyValuePair<string, string> pair in dicHeaderData)
                            for (int ikey = 1; ikey < arKeyName.Length; ikey++)
                            {

                                var rows = dicHeaderData.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                            == arKeyName[ikey].ToLower().Trim()).ToList();

                                if (arKeyName[ikey].Trim() == "FHTOTP" || arKeyName[ikey].Trim() == "FHSWGT")
                                {
                                    if (rows.Count > 0)
                                    {
                                        if (!string.IsNullOrEmpty(rows[0].Value))
                                        {
                                            // here you can use it safety
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = Convert.ToInt32(rows[0].Value.Trim());
                                        }
                                        else
                                        {

                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                        }
                                    }
                                    else
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                    }
                                }
                                else
                                {

                                    if (rows.Count > 0)
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Convert.ToString(rows[0].Value.Trim());
                                    }
                                    else
                                    {
                                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                    }
                                }

                            }

                            resultVal = command.ExecuteNonQuery();
                        }
                    }

                    //}
                    if (connection.State == ConnectionState.Open)
                    { connection.Close(); }
                }
                return resultVal;
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Insert FRP001--->" + ex,LogType.E);
                return resultVal;
            }
        }

        public static int UpdateIntoFRP001(Dictionary<string, string> dicHeaderData_FRP001, string connectionString, string ProNumber)
        {
            int resultVal = 0;
            string queryString = "";
            
            try
            {
                StringBuilder sbHeaderUpdate = new StringBuilder();
                StringBuilder sbHeader = new StringBuilder();
                //StringBuilder sbHeaderValue = new StringBuilder();
                sbHeader.Append("FHPRO");
                //sbHeaderValue.Append("?");
                foreach (KeyValuePair<string, string> pairkey in dicHeaderData_FRP001)
                {
                     sbHeader.Append("," + pairkey.Key.Trim());
                    //sbHeaderValue.Append(",?");
                    sbHeaderUpdate.Append(pairkey.Key.Trim() + "=? ,");
                }

                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    //if (dr != null)
                    //{

                    if (connection.State == ConnectionState.Closed)
                    {
                        connection.Open();
                    }

                    #region"SELECT"
                    //DataTable dtFHPRO = new DataTable();
                    //queryString = "select * from FRP001 where FHPRO=" + ProNumber;
                    //if (connection != null && !string.IsNullOrEmpty(queryString))
                    //{
                    //    using (OdbcCommand cmd = new OdbcCommand(queryString, connection))
                    //    {
                    //        cmd.CommandType = CommandType.Text;
                    //        using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                    //        {
                    //            da.Fill(dtFHPRO);
                    //        }
                    //    }
                    //}
                    #endregion 

                    //if (dtFHPRO != null && dtFHPRO.Rows.Count > 0)
                    //{
                        #region"UPDATE"
                        //queryString = "UPDATE FRP001 SET FDUID =? WHERE FDPRO =?  AND FDLIN =?";

                        queryString = "UPDATE FRP001 SET " + Convert.ToString(sbHeaderUpdate).Trim(',').Trim() + " WHERE FHPRO =? ";

                        // queryString = "INSERT INTO FRP001 (" + Convert.ToString(sbHeader).Trim() + ") VALUES (" + Convert.ToString(sbHeaderValue).Trim() + ")";
                        // queryString = "INSERT INTO FRP001 (FHPRO,FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
                        if (connection != null && !string.IsNullOrEmpty(queryString))
                        {
                            using (OdbcCommand command = new OdbcCommand(queryString, connection))
                            {

                                string[] arKeyName = Convert.ToString(sbHeader).Trim().Split(',');

                                for (int ikey = 1; ikey < arKeyName.Length; ikey++)
                                {

                                    var rows = dicHeaderData_FRP001.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                                == arKeyName[ikey].ToLower().Trim()).ToList();

                                    if (arKeyName[ikey].Trim() == "FHTOTP" || arKeyName[ikey].Trim() == "FHSWGT")
                                    {
                                        if (rows.Count > 0)
                                        {
                                            if (!string.IsNullOrEmpty(rows[0].Value))
                                            {
                                                // here you can use it safety
                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = Convert.ToInt32(rows[0].Value.Trim());
                                            }
                                            else
                                            {

                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                            }
                                        }
                                        else
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                        }
                                    }
                                    else
                                    {

                                        if (rows.Count > 0)
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Convert.ToString(rows[0].Value.Trim());
                                        }
                                        else
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                        }
                                    }

                                }

                                command.Parameters.Add("FHPRO", OdbcType.VarChar).Value = ProNumber.Trim();

                                resultVal = command.ExecuteNonQuery();
                            }
                        }

                        //}
                        #endregion
                    //}
                    //else
                    //{
                        #region "INSERT"
                        //queryString = "INSERT INTO FRP001 (" + Convert.ToString(sbHeader).Trim() + ") VALUES (" + Convert.ToString(sbHeaderValue).Trim() + ")";
                        //// queryString = "INSERT INTO FRP001 (FHPRO,FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
                        //if (connection != null && !string.IsNullOrEmpty(queryString))
                        //{
                        //    using (OdbcCommand command = new OdbcCommand(queryString, connection))
                        //    {
                        //        command.Parameters.Add("FHPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                        //        string[] arKeyName = Convert.ToString(sbHeader).Trim().Split(',');

                        //        // foreach (KeyValuePair<string, string> pair in dicHeaderData)
                        //        //foreach (KeyValuePair<string, string> pair in dicHeaderData)
                        //        for (int ikey = 1; ikey < arKeyName.Length; ikey++)
                        //        {

                        //            var rows = dicHeaderData.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                        //                                                        == arKeyName[ikey].ToLower().Trim()).ToList();

                        //            if (arKeyName[ikey].Trim() == "FHTOTP" || arKeyName[ikey].Trim() == "FHSWGT")
                        //            {
                        //                if (rows.Count > 0)
                        //                {
                        //                    if (!string.IsNullOrEmpty(rows[0].Value))
                        //                    {
                        //                        // here you can use it safety
                        //                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = Convert.ToInt32(rows[0].Value.Trim());
                        //                    }
                        //                    else
                        //                    {

                        //                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                        //                    }
                        //                }
                        //                else
                        //                {
                        //                    command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                        //                }
                        //            }
                        //            else
                        //            {

                        //                if (rows.Count > 0)
                        //                {
                        //                    command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Convert.ToString(rows[0].Value.Trim());
                        //                }
                        //                else
                        //                {
                        //                    command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                        //                }
                        //            }

                        //        }

                        //        resultVal = command.ExecuteNonQuery();
                        //    }
                        //}

                        ////}
                        #endregion
                    //}

                    if (connection.State == ConnectionState.Open)
                    { connection.Close();

                    }
                }
                return resultVal;
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in UpdateIntoFRP001()--->" + ex, LogType.E);
                return resultVal;
            }
        }

        public static int Insert_UpdateIntoFRP001(Dictionary<string, string> dicHeaderData, string connectionString, string ProNumber)
        {
            int resultVal = 0;
            string queryString = "";
            string HeaderValue = "";
            //string strKeyName = "FHPRO,FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC";
            //string strKeyName = "FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC";
            try
            {
                StringBuilder sbHeaderUpdate = new StringBuilder();
                StringBuilder sbHeader = new StringBuilder();
                StringBuilder sbHeaderValue = new StringBuilder();
                sbHeader.Append("FHPRO");
                sbHeaderValue.Append("?");
                foreach (KeyValuePair<string, string> pairkey in dicHeaderData)
                {
                    sbHeader.Append("," + pairkey.Key.Trim());
                    sbHeaderValue.Append(",?");
                    sbHeaderUpdate.Append(pairkey.Key.Trim() + "=? ,");
                }

                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    //if (dr != null)
                    //{

                    if (connection.State == ConnectionState.Closed)
                    {
                        connection.Open();
                    }

                    #region"SELECT"
                    DataTable dtFHPRO = new DataTable();
                    queryString = "select * from FRP001 where FHPRO=" + ProNumber;
                    if (connection != null && !string.IsNullOrEmpty(queryString))
                    {
                        using (OdbcCommand cmd = new OdbcCommand(queryString, connection))
                        {
                            cmd.CommandType = CommandType.Text;
                            using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                            {
                                da.Fill(dtFHPRO);
                            }
                        }
                    }
                    #endregion 

                    if (dtFHPRO != null && dtFHPRO.Rows.Count > 0)
                    {
                        #region"UPDATE"
                        //queryString = "UPDATE FRP001 SET FDUID =? WHERE FDPRO =?  AND FDLIN =?";

                        queryString = "UPDATE FRP001 SET " + Convert.ToString(sbHeaderUpdate).Trim(',').Trim() + " WHERE FHPRO =? ";

                        // queryString = "INSERT INTO FRP001 (" + Convert.ToString(sbHeader).Trim() + ") VALUES (" + Convert.ToString(sbHeaderValue).Trim() + ")";
                        // queryString = "INSERT INTO FRP001 (FHPRO,FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
                        if (connection != null && !string.IsNullOrEmpty(queryString))
                        {
                            using (OdbcCommand command = new OdbcCommand(queryString, connection))
                            {
                                
                                string[] arKeyName = Convert.ToString(sbHeader).Trim().Split(',');

                                for (int ikey = 1; ikey < arKeyName.Length; ikey++)
                                {

                                    var rows = dicHeaderData.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                                == arKeyName[ikey].ToLower().Trim()).ToList();

                                    if (arKeyName[ikey].Trim() == "FHTOTP" || arKeyName[ikey].Trim() == "FHSWGT")
                                    {
                                        if (rows.Count > 0)
                                        {
                                            if (!string.IsNullOrEmpty(rows[0].Value))
                                            {
                                                // here you can use it safety
                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = Convert.ToInt32(rows[0].Value.Trim());
                                            }
                                            else
                                            {

                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                            }
                                        }
                                        else
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                        }
                                    }
                                    else
                                    {

                                        if (rows.Count > 0)
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Convert.ToString(rows[0].Value.Trim());
                                        }
                                        else
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                        }
                                    }

                                }

                                command.Parameters.Add("FHPRO", OdbcType.VarChar).Value = ProNumber.Trim();

                                resultVal = command.ExecuteNonQuery();
                            }
                        }

                        //}
                        #endregion
                    }
                    else
                    {
                        #region "INSERT"
                        queryString = "INSERT INTO FRP001 (" + Convert.ToString(sbHeader).Trim() + ") VALUES (" + Convert.ToString(sbHeaderValue).Trim() + ")";
                        // queryString = "INSERT INTO FRP001 (FHPRO,FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
                        if (connection != null && !string.IsNullOrEmpty(queryString))
                        {
                            using (OdbcCommand command = new OdbcCommand(queryString, connection))
                            {
                                command.Parameters.Add("FHPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                                string[] arKeyName = Convert.ToString(sbHeader).Trim().Split(',');

                                // foreach (KeyValuePair<string, string> pair in dicHeaderData)
                                //foreach (KeyValuePair<string, string> pair in dicHeaderData)
                                for (int ikey = 1; ikey < arKeyName.Length; ikey++)
                                {

                                    var rows = dicHeaderData.AsEnumerable().Where(r => r.Key.ToLower().Trim()
                                                                                == arKeyName[ikey].ToLower().Trim()).ToList();

                                    if (arKeyName[ikey].Trim() == "FHTOTP" || arKeyName[ikey].Trim() == "FHSWGT")
                                    {
                                        if (rows.Count > 0)
                                        {
                                            if (!string.IsNullOrEmpty(rows[0].Value))
                                            {
                                                // here you can use it safety
                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = Convert.ToInt32(rows[0].Value.Trim());
                                            }
                                            else
                                            {

                                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                            }
                                        }
                                        else
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
                                        }
                                    }
                                    else
                                    {

                                        if (rows.Count > 0)
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Convert.ToString(rows[0].Value.Trim());
                                        }
                                        else
                                        {
                                            command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
                                        }
                                    }

                                }

                                resultVal = command.ExecuteNonQuery();
                            }
                        }

                        //}
                        #endregion
                    }

                    if (connection.State == ConnectionState.Open)
                    { connection.Close(); }
                }
                return resultVal;
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Insert FRP001--->" + ex,LogType.E);
                return resultVal;
            }
        }

        public static DataTable SelectPOBLSNFromFRP001(string connectionString, string ProNumber)
        {
            string queryString = "";
            DataTable dtPOBLSN = new DataTable();
            
            try
            {

                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {

                    if (connection.State == ConnectionState.Closed)
                    {
                        connection.Open();
                    }

                    #region"SELECT"
                    
                    queryString = "select FHBL, FHSNUM, FHPO from FRP001 where FHPRO=" + ProNumber;
                    if (connection != null && !string.IsNullOrEmpty(queryString))
                    {
                        using (OdbcCommand cmd = new OdbcCommand(queryString, connection))
                        {
                            cmd.CommandType = CommandType.Text;
                            using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                            {
                                da.Fill(dtPOBLSN);
                            }
                        }
                    }
                    #endregion 


                    if (connection.State == ConnectionState.Open)
                    { connection.Close(); }
                }

               
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Insert FRP001--->" + ex, LogType.E);
            }

            return dtPOBLSN;

        }

        public static string SelectPOBLSNFromFRP001(string FiledName, string connectionString, string ProNumber)
        {
            string queryString = "";
            string strFiledValue = string.Empty;

            try
            {

                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {

                    if (connection.State == ConnectionState.Closed)
                    {
                        connection.Open();
                    }

                    #region"SELECT"
                    DataTable dtPOBLSN = new DataTable();
                    queryString = "select " + FiledName + " from FRP001 where FHPRO=" + ProNumber;
                    if (connection != null && !string.IsNullOrEmpty(queryString))
                    {
                        using (OdbcCommand cmd = new OdbcCommand(queryString, connection))
                        {
                            cmd.CommandType = CommandType.Text;
                            using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                            {
                                da.Fill(dtPOBLSN);

                                if (dtPOBLSN != null && dtPOBLSN.Rows.Count > 0)
                                {

                                    strFiledValue = Convert.ToString(dtPOBLSN.Rows[0][FiledName]);
                                }
                            }
                        }
                    }
                    #endregion 


                    if (connection.State == ConnectionState.Open)
                    { connection.Close(); }
                }


            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Insert FRP001--->" + ex, LogType.E);
            }

            return strFiledValue;

        }

        public static int InsertIntoFRP002(DataRow dr, int Linenumber, string connectionString, string ProNumber)
        {
            int resultVal = 0;
            string queryString = "";
            try
            {

                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {

                    if (dr != null)
                    {

                        if (connection.State == ConnectionState.Closed)
                        { connection.Open(); }

                        queryString = "INSERT INTO FRP002TRU (FDPRO,FDLIN,FDPCS,FDCMCL,FDHAZ,FDPKGC,FDDES,FDWGT) VALUES (?,?,?,?,?,?,?,?)";
                        if (connection != null && !string.IsNullOrEmpty(queryString))
                        {
                            using (OdbcCommand command = new OdbcCommand(queryString, connection))
                            {
                                string Pices = Convert.ToString(dr["PCS"]).Trim();
                                string Weight = Convert.ToString(dr["Weight"]).Trim();

                                command.Parameters.Add("FDPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                                command.Parameters.Add("FDLIN", OdbcType.Int).Value = Linenumber;

                                if (!string.IsNullOrEmpty(Pices))
                                {
                                    // here you can use it safety
                                    command.Parameters.Add("FDPCS", OdbcType.Int).Value = Convert.ToInt32(Pices);
                                }
                                else
                                {

                                    command.Parameters.Add("FDPCS", OdbcType.Int).Value = 0;//DBNull.Value;
                                }
                                command.Parameters.Add("FDCMCL", OdbcType.VarChar).Value = Convert.ToString(dr["Code"]);
                                command.Parameters.Add("FDHAZ", OdbcType.VarChar).Value = Convert.ToString(dr["H"]);
                                command.Parameters.Add("FDPKGC", OdbcType.VarChar).Value = Convert.ToString(dr["P"]);
                                command.Parameters.Add("FDDES", OdbcType.VarChar).Value = Convert.ToString(dr["Description"]);
                                if (!string.IsNullOrEmpty(Weight))
                                {
                                    command.Parameters.Add("FDWGT", OdbcType.Int).Value = Convert.ToInt32(Weight);
                                }
                                else
                                {
                                    command.Parameters.Add("FDWGT", OdbcType.Int).Value = 0;//DBNull.Value;
                                }
                                //command.Parameters.Add("FDUID", OdbcType.VarChar).Value = Convert.ToString(cls_AppConstant.TruBot);
                                resultVal = command.ExecuteNonQuery();

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
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in InsertIntoFRP002()--->" + ex, LogType.E);
                return resultVal;
            }
        }


        public static int InsertIntoFRP002_UsingTable(DataTable dtFRP002TRU, int Linenumber, string connectionString, string ProNumber)
        {
            int resultVal = 0;
            string queryString = "";
            try
            {

                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    if (dtFRP002TRU != null && dtFRP002TRU.Rows.Count > 0)
                    {
                        //if (dr != null)
                        //{

                        if (connection.State == ConnectionState.Closed)
                        { connection.Open(); }

                        foreach (DataRow dr in dtFRP002TRU.Rows)
                        {
                            Linenumber = Linenumber + 1;
                            queryString = "INSERT INTO FRP002TRU (FDPRO,FDLIN,FDPCS,FDCMCL,FDHAZ,FDPKGC,FDDES,FDWGT) VALUES (?,?,?,?,?,?,?,?)";
                            if (connection != null && !string.IsNullOrEmpty(queryString))
                            {
                                using (OdbcCommand command = new OdbcCommand(queryString, connection))
                                {
                                    string Pices = Convert.ToString(dr["PCS"]).Trim();
                                    string Weight = Convert.ToString(dr["Weight"]).Trim();

                                    command.Parameters.Add("FDPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                                    command.Parameters.Add("FDLIN", OdbcType.Int).Value = Linenumber;

                                    if (!string.IsNullOrEmpty(Pices))
                                    {
                                        // here you can use it safety
                                        command.Parameters.Add("FDPCS", OdbcType.Int).Value = Convert.ToInt32(Pices);
                                    }
                                    else
                                    {

                                        command.Parameters.Add("FDPCS", OdbcType.Int).Value = 0;//DBNull.Value;
                                    }
                                    command.Parameters.Add("FDCMCL", OdbcType.VarChar).Value = Convert.ToString(dr["Code"]);
                                    command.Parameters.Add("FDHAZ", OdbcType.VarChar).Value = Convert.ToString(dr["H"]);
                                    command.Parameters.Add("FDPKGC", OdbcType.VarChar).Value = Convert.ToString(dr["P"]);
                                    command.Parameters.Add("FDDES", OdbcType.VarChar).Value = Convert.ToString(dr["Description"]);
                                    if (!string.IsNullOrEmpty(Weight))
                                    {
                                        command.Parameters.Add("FDWGT", OdbcType.Int).Value = Convert.ToInt32(Weight);
                                    }
                                    else
                                    {
                                        command.Parameters.Add("FDWGT", OdbcType.Int).Value = 0;//DBNull.Value;
                                    }
                                    //command.Parameters.Add("FDUID", OdbcType.VarChar).Value = Convert.ToString(cls_AppConstant.TruBot);
                                    resultVal = command.ExecuteNonQuery();

                                }
                            }
                        }

                        //}
                        if (connection.State == ConnectionState.Open)
                        { connection.Close(); }
                    }
                }
                return resultVal;
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in InsertIntoFRP002_UsingTable()--->" + ex, LogType.E);
                return resultVal;
            }
        }

        public static int UpdateIntoFRP002(List<int> _listLinenumber_FRP002TRU, string connectionString, string ProNumber, OdbcParameter[] parameters = null)
        {
            int resultVal = 0;
            try
            {

                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    if (connection.State == ConnectionState.Closed)
                    { connection.Open(); }

                    foreach (var _itemlinenumber in _listLinenumber_FRP002TRU)
                    {
                        string queryString = "UPDATE FRP002TRU SET FDUID =? WHERE FDPRO =?  AND FDLIN =?";

                        if (connection != null && !string.IsNullOrEmpty(queryString))
                        {
                            using (OdbcCommand command = new OdbcCommand(queryString, connection))
                            {
                                command.CommandType = CommandType.Text;
                                command.Parameters.Add("FDUID", OdbcType.VarChar).Value = "TRUBOT";
                                command.Parameters.Add("FDPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                                command.Parameters.Add("FDLIN", OdbcType.Int).Value = _itemlinenumber;

                                resultVal = command.ExecuteNonQuery();

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
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in UpdateIntoFRP002--->" + ex,LogType.E);
                return 0;
            }
        }


        //public static int UpdateIntoFRPD010_UpdateIntoFRP001_InsertIntoFRP002(DataTable dtFRPD010, Dictionary<string, string> dicHeaderData_FRP001, DataTable dtFRP002TRU, int Linenumber, List<int> _listLinenumber_FRP002TRU, string connectionString, string ProNumber, OdbcParameter[] parameters = null)
        //{
        //    int resultVal = 0;
        //    string queryString = "";

        //    try
        //    {







        //        StringBuilder sbHeaderUpdate = new StringBuilder();
        //        StringBuilder sbHeader = new StringBuilder();
        //        sbHeader.Append("FHPRO");
        //        foreach (KeyValuePair<string, string> pairkey in dicHeaderData_FRP001)
        //        {
        //            sbHeader.Append("," + pairkey.Key.Trim());
        //            //sbHeaderValue.Append(",?");
        //            sbHeaderUpdate.Append(pairkey.Key.Trim() + "=? ,");
        //        }

        //        using (OdbcConnection connection = new OdbcConnection(connectionString))
        //        {
        //            //if (dr != null)
        //            //{

        //            if (connection.State == ConnectionState.Closed)
        //            {
        //                connection.Open();
        //            }

        //            #region"SELECT"
        //            //DataTable dtFHPRO = new DataTable();
        //            //queryString = "select * from FRP001 where FHPRO=" + ProNumber;
        //            //if (connection != null && !string.IsNullOrEmpty(queryString))
        //            //{
        //            //    using (OdbcCommand cmd = new OdbcCommand(queryString, connection))
        //            //    {
        //            //        cmd.CommandType = CommandType.Text;
        //            //        using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
        //            //        {
        //            //            da.Fill(dtFHPRO);
        //            //        }
        //            //    }
        //            //}
        //            #endregion

        //            //if (dtFHPRO != null && dtFHPRO.Rows.Count > 0)
        //            //{
        //            #region"UPDATE"
        //            //queryString = "UPDATE FRP001 SET FDUID =? WHERE FDPRO =?  AND FDLIN =?";

        //            queryString = "UPDATE FRP001 SET " + Convert.ToString(sbHeaderUpdate).Trim(',').Trim() + " WHERE FHPRO =? ";

        //            // queryString = "INSERT INTO FRP001 (" + Convert.ToString(sbHeader).Trim() + ") VALUES (" + Convert.ToString(sbHeaderValue).Trim() + ")";
        //            // queryString = "INSERT INTO FRP001 (FHPRO,FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
        //            if (connection != null && !string.IsNullOrEmpty(queryString))
        //            {
        //                using (OdbcCommand command = new OdbcCommand(queryString, connection))
        //                {

        //                    string[] arKeyName = Convert.ToString(sbHeader).Trim().Split(',');

        //                    for (int ikey = 1; ikey < arKeyName.Length; ikey++)
        //                    {

        //                        var rows = dicHeaderData_FRP001.AsEnumerable().Where(r => r.Key.ToLower().Trim()
        //                                                                    == arKeyName[ikey].ToLower().Trim()).ToList();

        //                        if (arKeyName[ikey].Trim() == "FHTOTP" || arKeyName[ikey].Trim() == "FHSWGT")
        //                        {
        //                            if (rows.Count > 0)
        //                            {
        //                                if (!string.IsNullOrEmpty(rows[0].Value))
        //                                {
        //                                    // here you can use it safety
        //                                    command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = Convert.ToInt32(rows[0].Value.Trim());
        //                                }
        //                                else
        //                                {

        //                                    command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
        //                                }
        //                            }
        //                            else
        //                            {
        //                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
        //                            }
        //                        }
        //                        else
        //                        {

        //                            if (rows.Count > 0)
        //                            {
        //                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Convert.ToString(rows[0].Value.Trim());
        //                            }
        //                            else
        //                            {
        //                                command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
        //                            }
        //                        }

        //                    }

        //                    command.Parameters.Add("FHPRO", OdbcType.VarChar).Value = ProNumber.Trim();

        //                    resultVal = command.ExecuteNonQuery();
        //                }
        //            }

        //            //}
        //            #endregion
        //            //}
        //            //else
        //            //{
        //            #region "INSERT"
        //            //queryString = "INSERT INTO FRP001 (" + Convert.ToString(sbHeader).Trim() + ") VALUES (" + Convert.ToString(sbHeaderValue).Trim() + ")";
        //            //// queryString = "INSERT INTO FRP001 (FHPRO,FHBL,FHSNUM,FHPO,FHSCD,FHSNM,FHSA1,FHSA2,FHSCT,FHSST,FHSZIP,FHCCD,FHCNM,FHCA1,FHCA2,FHCCT,FHCST,FHCZIP,FHBTC,FHBNM,FHBA1,FHBA2,FHBCT,FHBST,FHBZIP,FHTRM,FHTOTP,FHSWGT,FHRADF,FHRADT,FHCOD,FHDC) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
        //            //if (connection != null && !string.IsNullOrEmpty(queryString))
        //            //{
        //            //    using (OdbcCommand command = new OdbcCommand(queryString, connection))
        //            //    {
        //            //        command.Parameters.Add("FHPRO", OdbcType.VarChar).Value = ProNumber.Trim();
        //            //        string[] arKeyName = Convert.ToString(sbHeader).Trim().Split(',');

        //            //        // foreach (KeyValuePair<string, string> pair in dicHeaderData)
        //            //        //foreach (KeyValuePair<string, string> pair in dicHeaderData)
        //            //        for (int ikey = 1; ikey < arKeyName.Length; ikey++)
        //            //        {

        //            //            var rows = dicHeaderData.AsEnumerable().Where(r => r.Key.ToLower().Trim()
        //            //                                                        == arKeyName[ikey].ToLower().Trim()).ToList();

        //            //            if (arKeyName[ikey].Trim() == "FHTOTP" || arKeyName[ikey].Trim() == "FHSWGT")
        //            //            {
        //            //                if (rows.Count > 0)
        //            //                {
        //            //                    if (!string.IsNullOrEmpty(rows[0].Value))
        //            //                    {
        //            //                        // here you can use it safety
        //            //                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = Convert.ToInt32(rows[0].Value.Trim());
        //            //                    }
        //            //                    else
        //            //                    {

        //            //                        command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
        //            //                    }
        //            //                }
        //            //                else
        //            //                {
        //            //                    command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.Int).Value = 0;//DBNull.Value;
        //            //                }
        //            //            }
        //            //            else
        //            //            {

        //            //                if (rows.Count > 0)
        //            //                {
        //            //                    command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = Convert.ToString(rows[0].Value.Trim());
        //            //                }
        //            //                else
        //            //                {
        //            //                    command.Parameters.Add(arKeyName[ikey].Trim(), OdbcType.VarChar).Value = "";
        //            //                }
        //            //            }

        //            //        }

        //            //        resultVal = command.ExecuteNonQuery();
        //            //    }
        //            //}

        //            ////}
        //            #endregion
        //            //}

        //            if (connection.State == ConnectionState.Open)
        //            {
        //                connection.Close();

        //            }
        //        }
        //        return resultVal;
        //    }
        //    catch (Exception ex)
        //    {
        //        BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in UpdateIntoFRP001()--->" + ex, LogType.E);
        //        return resultVal;
        //    }
        //}
        public static int GetDataTableByTextInsert(List<int> _listLinenumber, string connectionString, string ProNumber, OdbcParameter[] parameters = null)
        {
            int resultVal = 0;
            try
            {

                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    if (connection.State == ConnectionState.Closed)
                    { connection.Open(); }
                    // string queryString = "UPDATE FRP002 SET FDUID = 'TRUBOT' WHERE FDPRO = '10471180680' and FDLIN IN(481,482,483)";
                    //  string  queryString = "UPDATE FRP002 SET FDUID='" + cls_AppConstant.TruBot.Trim() + "' WHERE FDPRO = '" + ProNumber.Trim() + "' and FDLIN=" + Linenumber + "";

                    foreach (var _itemlinenumber in _listLinenumber)
                    {
                        string queryString = "UPDATE FRP002 SET FDUID =? WHERE FDPRO =?  AND FDLIN =?";

                        if (connection != null && !string.IsNullOrEmpty(queryString))
                        {
                            using (OdbcCommand command = new OdbcCommand(queryString, connection))
                            {
                                command.CommandType = CommandType.Text;
                                command.Parameters.Add("FDUID", OdbcType.VarChar).Value = "TRUBOT";
                                command.Parameters.Add("FDPRO", OdbcType.VarChar).Value = ProNumber.Trim();
                                command.Parameters.Add("FDLIN", OdbcType.Int).Value = _itemlinenumber;

                                resultVal = command.ExecuteNonQuery();

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
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in GetDataTableByTextInsert--->" + ex,LogType.E );
                return 0;
            }
        }
        public static int GetLineNumberFromFRP002(string connectionString, string ProNumber)
        {
            int resultVal = 0;
            string queryString = "";
            DataTable dt = new DataTable();
            try
            {

                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == false)
                {
                    OdbcConnection db2Connection = new OdbcConnection();
                    //MessageBox.Show("GetDataTableByText_IsEDI: DB2ConnectionString: " + DB2ConnectionString);
                    db2Connection.ConnectionString = DBUtilities.DB2ConnectionString;

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
                    connection.ConnectionString = DBUtilities.DB2ConnectionString;
                    connection.Open();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("DB2Conn");
                    }
                    iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", connection);
                }



                //using (OdbcConnection connection = new OdbcConnection(connectionString))
                //{
                    //if (connection.State == ConnectionState.Closed)
                    //{ connection.Open(); }

                    queryString = "SELECT MAX(FDLIN) as LineNumber FROM FRP002 WHERE FDPRO='" + ProNumber + "'";
                    if (connection != null && !string.IsNullOrEmpty(queryString))
                    {
                        using (OdbcCommand command = new OdbcCommand(queryString, connection))
                        {
                            OdbcDataAdapter sda = new OdbcDataAdapter(queryString, connection);
                            sda.Fill(dt);
                            if (dt.Rows.Count > 0)
                            {
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("Databe count--->" + dt.Rows.Count, LogType.P );


                                string LineVaue = Convert.ToString(dt.Rows[0]["LineNumber"]);
                                resultVal = Convert.ToInt32(LineVaue);
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("Linenumber-->" + resultVal, LogType.P);
                            }
                        }
                    }


                    //if (connection.State == ConnectionState.Open)
                    //{ connection.Close(); }


                //}
                return resultVal;
            }
            catch (Exception ex)
            {
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in GetLineNumberFromFRP002()--->" + ex, LogType.E);
                return resultVal;
            }
        }

        public static int GetLineNumberFromFRP002_BackUp(string connectionString, string ProNumber)
        {
            int resultVal = 0;
            string queryString = "";
            DataTable dt = new DataTable();
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    if (connection.State == ConnectionState.Closed)
                    { connection.Open(); }

                    queryString = "SELECT MAX(FDLIN) as LineNumber FROM FRP002 WHERE FDPRO='" + ProNumber + "'";
                    if (connection != null && !string.IsNullOrEmpty(queryString))
                    {
                        using (OdbcCommand command = new OdbcCommand(queryString, connection))
                        {
                            OdbcDataAdapter sda = new OdbcDataAdapter(queryString, connection);
                            sda.Fill(dt);
                            if (dt.Rows.Count > 0)
                            {
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("Databe count--->" + dt.Rows.Count, LogType.P);


                                string LineVaue = Convert.ToString(dt.Rows[0]["LineNumber"]);
                                resultVal = Convert.ToInt32(LineVaue);
                                BOLDocumentDataExtraction.DataExtraction.WriteLog("Linenumber-->" + resultVal, LogType.P);
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
                BOLDocumentDataExtraction.DataExtraction.WriteLog("Exception in GetLineNumberFromFRP002()--->" + ex, LogType.E);
                return resultVal;
            }
        }

        #endregion
    }
}
