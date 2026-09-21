using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;
using System.IO;

namespace BOLSET1IQ
{
    public class clsF28
    {

        public RetStructF28 F28(string strUserName, string strIQMode, string strLocation, string strRoutineNo, string strArg)
        {
            RetStructF28 objStructF28 = new RetStructF28();
            objStructF28.Status = "F";

            try
            {
                switch (strRoutineNo)
                {
                    case "C75":

                        objStructF28 = RetrieveTiff(strUserName);

                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("F28 - " + ex.Message.ToString());
            }
            return objStructF28;
        }


        public static RetStructF28 RetrieveTiff(string UserName)
        {
            RetStructF28 stc28 = new RetStructF28();
            stc28.Status = "F";


            string strFileName = string.Empty;
            string strTiffPath = string.Empty;
            string sBatchname = string.Empty;

            string TiffImagesPath = @"D:\E Drive\G DRIVE\2020\Shantilal\Demo Images\Demo Images_tiff";
           
            try
            {
                

                DirectoryInfo objDirecinfo = new DirectoryInfo(TiffImagesPath);

                FileInfo[] objFileInfo = objDirecinfo.GetFiles("*.tif");

                if (objFileInfo.Length > 0)
                {
                    string strtifffullpath = string.Empty;
                    for (int i = 0; i <= objFileInfo.Length - 1; i++)
                    {
                        strtifffullpath = objFileInfo[i].FullName;
                        if (File.Exists(objFileInfo[i].DirectoryName  + @"\" + objFileInfo[i].Name.Replace(objFileInfo[i].Extension, ".xml")) == false)
                        {
                            if (File.Exists(objFileInfo[i].DirectoryName + @"\" + objFileInfo[i].Name.Replace(objFileInfo[i].Extension, ".pro")) == true)
                            {

                                stc28.Status = "S";
                                stc28.inputTiffPath = strtifffullpath;
                                stc28.XmlFilePath = objFileInfo[i].DirectoryName + @"\" + objFileInfo[i].Name.Replace(objFileInfo[i].Extension, ".xml");
                                stc28.BatchName = "Demo Images_tiff";
                                break;
                            }
                        }
                    }
                }
            }



            

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                stc28.Status = "F";
            }


            return stc28;
        }

    }
}
