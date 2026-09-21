using System;

using System.Collections.Generic;
using System.Linq;
using System.Text;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;
//using BOL_SET1;

namespace BOLDocumentDataExtraction.BLL
{
    
    public class DataExtraction
    {

        #region "Not Required"

        //public static Dictionary<string, string> FieldList = new Dictionary<string, string>();
        //public static string Synonyms = string.Empty;

        //private static StringBuilder AppendExtractedData(RetStructF3 possibleWord, string FieldName, long documentID)
        //{
        //    StringBuilder objExtractedData = new StringBuilder();

        //    objExtractedData.Append("<Header>");
        //    objExtractedData.Append("<DocumentID> " + documentID + " </DocumentID>");
        //    objExtractedData.Append("<FieldName> " + FieldName + " </FieldName>");

        //    if (possibleWord.Status == "S")
        //    {
        //        objExtractedData.Append("<FieldValue>" + possibleWord.Words[0].strWord + "</FieldValue>");
        //        objExtractedData.Append("<AvgOCRConfLevel>" + possibleWord.Words[0].Confidence + "</AvgOCRConfLevel>");
        //        objExtractedData.Append("<CharConflevel>" + possibleWord.Words[0].ConfString + " </CharConflevel>");
        //        objExtractedData.Append("<SearchConfLevel>" + possibleWord.ConfidenceLevelofSuspect[0] + "</SearchConfLevel>");
        //        objExtractedData.Append("<Remarks> AutoCaptured </Remarks>");
        //        objExtractedData.Append("<PageNO>" + possibleWord.Words[0].PageNo + "</PageNO>");
        //        objExtractedData.Append("<ROI>" + possibleWord.Words[0].Left + ',' + possibleWord.Words[0].Top + ',' + +possibleWord.Words[0].Right + ',' + possibleWord.Words[0].Bottom + "</ROI>");
        //    }
        //    else
        //    {
        //        objExtractedData.Append("<FieldValue>  </FieldValue>");
        //        objExtractedData.Append("<AvgOCRConfLevel>  </AvgOCRConfLevel>");
        //        objExtractedData.Append("<CharConflevel> </CharConflevel>");
        //        objExtractedData.Append("<SearchConfLevel>  </SearchConfLevel>");
        //        objExtractedData.Append("<Remarks>  </Remarks>");
        //        objExtractedData.Append("<PageNO>  </PageNO>");
        //        objExtractedData.Append("<ROI>  </ROI>");
        //    }

        //    objExtractedData.Append("</Header>");

        //    return objExtractedData;

        //}
        //public static string DataExtraction_BOL(clsCncMetaData ObjMetaData, int intCurrPageNumber, long documentID)
        //{
        //    string DataExtraction_BOL = string.Empty;
        //    StringBuilder objDataExtraction = new StringBuilder();

        //    try
        //    {

        //        string xmlversion = "<?xml version=" + "'1.0'" + "?>";

        //        objDataExtraction.Append(xmlversion);
        //        objDataExtraction.Append("<HeaderDetails>");

        //        foreach (KeyValuePair<string, string> fieldName in FieldList)
        //        {
        //            RetStructF3 retStructValue = DataExtraction_F3.F3(ObjMetaData, intCurrPageNumber, "", fieldName.Value);
        //            if (retStructValue.Status == "S")
        //            {
        //                retStructValue.Words[0].strWord = DataExtraction_F5.F5(retStructValue.Words[0].strWord.Trim(), "", fieldName.Value);
        //                //if (DataExtraction_F7.F7(retStructValue.Words[0].strWord.Trim(), fieldName.Value) == false)
        //                //{
        //                //    ReturnPurchaseOrder = "";
        //                //}
        //            }

        //            objDataExtraction.Append(AppendExtractedData(retStructValue, fieldName.Key, documentID));
        //        }

        //        objDataExtraction.Append("</HeaderDetails>");

        //        DataExtraction_BOL = Convert.ToString(objDataExtraction);

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("DataExtraction_BOL: " + ex.Message);

        //    }

        //    return DataExtraction_BOL;
        //}

        #endregion


        //static BOL_SET1.clsF3 objBOLF3 = new clsF3();
        //private static StringBuilder AppendExtractedData1(string FieldName, string FieldValue, float AvgOCRConfLevel, float CharConflevel, float SearchConfLevel, string Remarks, int PageNO, string ROI)
        //{
        //    StringBuilder objDataExtraction = new StringBuilder();
        //    //objDataExtraction.Append("< DocumentID > " + documentID + " </ DocumentID >");
        //    //objDataExtraction.Append("< DocumentID > " + documentID + " </ DocumentID >");
        //    //objDataExtraction.Append("< DocumentID > " + documentID + " </ DocumentID >");
        //    //objDataExtraction.Append("< DocumentID > " + documentID + " </ DocumentID >");
        //    //objDataExtraction.Append("< DocumentID > " + documentID + " </ DocumentID >");
        //    //objDataExtraction.Append("< DocumentID > " + documentID + " </ DocumentID >");

        //    return objDataExtraction;

        //}
        //private static StringBuilder AppendBlankData(string FieldName, string Remarks)
        //{
        //    StringBuilder objBlankData = new StringBuilder();


        //    objBlankData.Append("< FieldName > " + FieldName + " </ FieldName >");
        //    objBlankData.Append("< FieldValue >  </ FieldValue >");
        //    objBlankData.Append("< AvgOCRConfLevel >  </ AvgOCRConfLevel >");
        //    objBlankData.Append("< CharConflevel > </ CharConflevel >");
        //    objBlankData.Append("< SearchConfLevel >  </ SearchConfLevel >");
        //    objBlankData.Append("< Remarks > " + Remarks + " </ Remarks >");
        //    objBlankData.Append("< PageNO >  </ SearchConfLevel >");
        //    objBlankData.Append("< ROI >  </ Remarks >");

        //    return objBlankData;

        //}



    }
}
