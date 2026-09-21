using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Text;

namespace BOLSET1IQ
{
    public class clsF25
    {
        DirectoryInfo opFolder;
        FileInfo[] infoFile;
        private XmlDocument xmlDoc = new XmlDocument();

        public RetStructF25 F25(string strFileName, string strRoutineNo, string strArg)
        {
            RetStructF25 objStructF25 = new RetStructF25();
            objStructF25.Status = "F";

            string filepath = strFileName.Substring(0, strFileName.LastIndexOf(@"\"));
            string[] files = System.IO.Directory.GetFiles(filepath, "*.xml");

            try
            {
                switch (strRoutineNo)
                {
                    case "C75":

                        DataTable dt = new DataTable();

                        dt.Columns.Add("Tiff Name");
                        dt.Columns.Add("BOL");
                        dt.Columns.Add("PO");
                        dt.Columns.Add("SN");
                        dt.Columns.Add("HazMat Telephone");
                        dt.Columns.Add("Indivdual Telephone");                        

                        for (int i = 0; i < files.Length; i++)
                        {
                            xmlDoc.Load(files[i]);

                            XmlNodeList xmlFieldNodes = xmlDoc.SelectNodes("DocumentInfo//FieldInfo");
                            int col = 0;
                            DataRow dr = dt.NewRow();
                            int x = 0;
                            //foreach (XmlNode xmlFieldNode in xmlFieldNodes)
                            //{
                            while (x < xmlFieldNodes.Count)
                            {

                                try
                                {
                                    if (col < dt.Columns.Count)
                                    {
                                        if (col == 0)
                                        {
                                            dr[col] = xmlFieldNodes[x].Attributes["TIFFNAME"].Value;
                                            col++;
                                            x++;
                                        }
                                        //if (col == 1)
                                        //{
                                        //    dr[col] = xmlFieldNodes[1].Attributes["fldval"].Value;
                                        //    x++;
                                        //    col++;
                                        //}
                                        else
                                        {
                                            dr[col] = xmlFieldNodes[x].Attributes["fldval"].Value;
                                            x++;
                                            col++;
                                        }

                                    }

                                    else { break; }
                                    //dt.Rows[i][col] = xmlFieldNode.Attributes["fldval"].Value;
                                    //strb1.AppendLine(xmlFieldNode.Attributes["fldname"].Value + ":" + xmlFieldNode.Attributes["fldval"].Value);


                                    //strb.Append(xmlFieldNode.Attributes["fldval"].Value);                                    
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.Message);
                                }

                            }
                            dt.Rows.Add(dr);
                        }

                        //ExportExcel(app, wb, dt);
                        ExportToCSV(dt);

                        //read(strFileName);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("F25 - " + ex.Message.ToString());
            }
            return objStructF25;
        }

        private static void ExportToCSV(DataTable dt)
        {
            string OutputFolder = Application.StartupPath + "\\" + "Output";
            string SavePath = OutputFolder + "\\" + "Output_CSVFile" + ".csv";
            StringBuilder sb = new StringBuilder();
            foreach (DataColumn dc in dt.Columns)
            {
                sb.Append(dc.ColumnName + ",");
            }
            sb.Remove(sb.Length - 1, 1);
            sb.AppendLine();
            foreach (DataRow row in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sb.Append(row[i].ToString() + ",");
                }
                sb.Remove(sb.Length - 1, 1);
                sb.AppendLine();
            }

            File.WriteAllText(SavePath, sb.ToString());
        }

        private string read(string strFileName)
        {

            XmlDocument objdoc = new XmlDocument();

            opFolder = new DirectoryInfo(Path.GetDirectoryName(strFileName));

            infoFile = opFolder.GetFiles("*.xml");

            for (int i = 0; i <= infoFile.Length; i++)
            {
                objdoc.Load(infoFile[i].FullName);

                XmlNode objnode = objdoc.SelectSingleNode("DocumentInfo//FieldInfo");

                string filedvale = objnode.Attributes["fldval"].Value;

                string filedvale1 = objdoc.SelectSingleNode("DocumentInfo//FieldInfo[@fieldid=2]").Attributes["fldval"].Value.Trim();
                
                MessageBox.Show(filedvale1);

            }


            return "";
        }
    }
}