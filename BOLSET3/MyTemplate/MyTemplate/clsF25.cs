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

namespace BOLSET3IQ
{
    public class clsF25
    {
        DirectoryInfo opFolder;
        FileInfo[] infoFile;


        public RetStructF25 F25(string strFileName, string strRoutineNo, string strArg)
        {
            RetStructF25 objStructF25 = new RetStructF25();
            objStructF25.Status = "F";

            try
            {
                switch (strRoutineNo)
                {
                    case "C75":

                        read(strFileName);
                        break;
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show("F25 - " + ex.Message.ToString());
            }
            return objStructF25;
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
                
               // MessageBox.Show(filedvale1);

            }


            return "";
        }
    }
}