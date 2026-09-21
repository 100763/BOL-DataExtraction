using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Data;

namespace BOLDocumentDataExtraction.CommonModule
{
   public class DataExchange
    {
        public static DataTable dtDataExg;
        public static int int_srno;

        enum FileType
        {
            XML = 1,
            Text = 2,
            CSV = 3,
            Excel = 4
        }

        //public DataExchange()
        //{
        //    CreateStructure();
        //}

        public static void CreateStructure()
        {
            try
            {
                dtDataExg = new DataTable("DataExchange");
                dtDataExg.Columns.Add("SrNo", System.Type.GetType("System.String"));
                dtDataExg.Columns.Add("TiffName", System.Type.GetType("System.String"));
                dtDataExg.Columns.Add("PageNo", System.Type.GetType("System.String"));
                dtDataExg.Columns.Add("FieldNo", System.Type.GetType("System.String"));
                dtDataExg.Columns.Add("FieldNm", System.Type.GetType("System.String"));
                dtDataExg.Columns.Add("FieldLvl", System.Type.GetType("System.String"));
                dtDataExg.Columns.Add("RoutineNo", System.Type.GetType("System.String"));
                dtDataExg.Columns.Add("AlgoNo", System.Type.GetType("System.String"));
                dtDataExg.Columns.Add("AlgoArgs", System.Type.GetType("System.String"));
                dtDataExg.Columns.Add("AlgoStatus", System.Type.GetType("System.String"));

                dtDataExg.Columns.Add("StartTime", System.Type.GetType("System.String"));
                dtDataExg.Columns.Add("EndTime", System.Type.GetType("System.String"));

                dtDataExg.Columns.Add("RetStr", System.Type.GetType("System.Object"));

                dtDataExg.Columns.Add("DocStartTime", System.Type.GetType("System.String"));
                dtDataExg.Columns.Add("DocEndTime", System.Type.GetType("System.String"));
                dtDataExg.DefaultView.Sort = "";
            }
            catch (Exception ex)
            {
                //Interaction.MsgBox(ex.Message);
            }
        }

        public static void UpdateRecBeforeCall(string strTiffName, string strFldNo, string strFieldNm, string strFldLvl, string strFunction, string strAlgo, string strArgs, Int32 intPageNo)
        {
            try
            {
                DataRow[] drTemp = dtDataExg.Select("TiffName='" + strTiffName.ToUpper() + "'");
                DataRow dr = dtDataExg.NewRow();
                int_srno += 1;
                dr["SrNo"] = Convert.ToString(int_srno);
                dr["TiffName"] = strTiffName.ToUpper();
                dr["PageNo"] = intPageNo.ToString();
                dr["FieldNo"] = strFldNo;
                dr["FieldNm"] = strFieldNm;
                dr["FieldLvl"] = strFldLvl;
                dr["RoutineNo"] = strFunction;
                dr["AlgoNo"] = strAlgo;
                dr["AlgoArgs"] = strArgs;
                dr["StartTime"] = DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Second.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Millisecond.ToString().PadLeft(3, '0');

                if ((drTemp.Length == 0))
                    dr["DocStartTime"] = DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Second.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Millisecond.ToString().PadLeft(3, '0');

                dtDataExg.Rows.Add(dr);
            }
            catch (Exception ex)
            {
                //Interaction.MsgBox(ex.Message);
            }
        }

        public static void UpdateRecAfterCall(string strTiffName, string strFldNo, string strFldLvl, string strFunction, string strAlgo, string algoStatus, object objData)
        {
            try
            {
                DataRow[] drTemp = dtDataExg.Select("tiffname = '" + strTiffName + "' and FieldNo='" + strFldNo + "' and FieldLvl = '" + strFldLvl + "' and RoutineNo = '" + strFunction + "' and AlgoNo = '" + strAlgo + "'");
               // DataRow[] drTemp = dtDataExg.Select("tiffname = '" + strTiffName + "' and FieldNo='" + strFldNo + "' and RoutineNo = '" + strFunction + "' and AlgoNo = '" + strAlgo + "'");
                for (int i = 0; i <= drTemp.Length - 1; i++)
                {

                    if (drTemp[i]["EndTime"] == DBNull.Value)
                    {
                        drTemp[i]["AlgoStatus"] = algoStatus;
                        drTemp[i]["EndTime"] = DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Second.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Millisecond.ToString().PadLeft(3, '0');
                        drTemp[i]["RetStr"] = objData;
                    }
                    else
                    {
                        continue;
                    }
                }
            }

            catch (Exception ex)
            {
                //Interaction.MsgBox(ex.Message);
            }
        }

        public static void ClearData()
        {
            try
            {
                dtDataExg.Rows.Clear();
            }
            catch (Exception ex)
            {
                //Interaction.MsgBox(ex.Message);
            }
        }

        public static void UpdateDocEndTime(string strTiffName)
        {
            DataRow[] drTemp = dtDataExg.Select("TiffName='" + strTiffName.ToUpper() + "'");
            if (drTemp.Length > 0)
                drTemp[0]["DocEndTime"] = DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Minute.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Second.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Millisecond.ToString().PadLeft(3, '0');
        }

        public static void WriteDatatoXML(string strFile)
        {
            try
            {
                DataTable dtResult = new DataTable();
                dtResult = dtDataExg.Copy();
                dtResult.Columns.Remove("RetStr");


                DataSet dsTemp = new DataSet();

                if (System.IO.File.Exists(strFile))
                {
                    dsTemp.ReadXml(strFile);

                    if (dsTemp.Tables.Count > 0)
                        dtResult.Merge(dsTemp.Tables[0]);
                }

                // write the data in xml file
                if ((System.IO.File.Exists(strFile)))
                    System.IO.File.Delete(strFile);

                dtResult.WriteXml(strFile);
                dtDataExg.Rows.Clear();
            }
            catch (Exception ex)
            {
               // Interaction.MsgBox("Error in MergeDataTable " + ex.Message);
            }
        }

        private static void AddColumnsToStruct()
        {
            try
            {
                for (int intCount = 1; intCount <= 27; intCount++)
                {
                    dtDataExg.Columns.Add("F" + intCount.ToString() + "Algo", System.Type.GetType("System.String"));
                    dtDataExg.Columns.Add("F" + intCount.ToString() + "Args", System.Type.GetType("System.String"));
                    dtDataExg.Columns.Add("F" + intCount.ToString() + "Ret", System.Type.GetType("System.String"));
                    dtDataExg.Columns.Add("F" + intCount.ToString() + "ST", System.Type.GetType("System.String"));
                    dtDataExg.Columns.Add("F" + intCount.ToString() + "ET", System.Type.GetType("System.String"));
                }
            }
            catch (Exception ex)
            {
                //Interaction.MsgBox(ex.Message);
            }
        }

    }
}
