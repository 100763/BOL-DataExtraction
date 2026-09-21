using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOLDocumentDataExtraction.CommonModule
{
    public static class Common
    {
        public static void DisposeProcessLog(this StreamWriter processLogWriter)
        {
            if (processLogWriter != null)
            {
                processLogWriter.Close();
                processLogWriter.Dispose();
            }
        }
        public static StreamWriter CreateProcessLogFile(string processLogFilePath, string processLogfileName)
        {
            StreamWriter processLogWriter = null;
            List<string> dateList = new List<string>();
            try
            {
                string logFolderName = ConfigurationManager.AppSettings["LogFolderName"].ToString();
                if (!processLogFilePath.EndsWith("\\"))
                {
                    processLogFilePath = processLogFilePath + "\\";
                }
                dateList.Add(processLogFilePath + logFolderName + "\\ProcessLog\\" + DateTime.Now.ToString("MM-dd-yyyy"));
                for (int i = 1; i <= 4; i++)
                {
                    dateList.Add(processLogFilePath + logFolderName + "\\ProcessLog\\" + DateTime.Now.AddDays(-i).ToString("MM-dd-yyyy"));
                }

                if (!Directory.Exists(processLogFilePath + logFolderName + "\\ProcessLog\\" + DateTime.Now.ToString("MM-dd-yyyy")))
                {
                    Directory.CreateDirectory(processLogFilePath + logFolderName + "\\ProcessLog\\" + DateTime.Now.ToString("MM-dd-yyyy"));
                }
                //processLogFilePath = AppDomain.CurrentDomain.BaseDirectory + "\\ProcessLog\\" + DateTime.Now.ToString("dd-MM-yyyy");
                processLogWriter = new StreamWriter(processLogFilePath + logFolderName + "\\ProcessLog\\" + DateTime.Now.ToString("MM-dd-yyyy") + "\\" + processLogfileName + "_" + DateTime.Now.ToString("HH-mm-ss") + ".txt");
                string[] subdirectoryEntries = Directory.GetDirectories(processLogFilePath + logFolderName + "\\ProcessLog\\");
                foreach (string subdirectory in subdirectoryEntries)
                {
                    if (!dateList.Contains(subdirectory))
                    {
                        Directory.Delete(subdirectory, true);
                    }
                }
            }
            catch (Exception)
            {

            }
            return processLogWriter;
        }
        public static string CalculateTotalTime(DateTime startTime,DateTime endTime)
        {
            TimeSpan ts = endTime - startTime;
            string message = " " + ts.Minutes + " : " + ts.Seconds;
            return message;
        }


        //public static string CalculateTotalTime(DateTime startTime, DateTime endTime)
        //{
        //    TimeSpan ts = endTime - startTime;
        //    string message = "Hours:" + ts.Hours + " Minutes: " + ts.Minutes + "Seconds:" + ts.Seconds;
        //    return message;
        //}
        public static string CalculateTotalTime_Sec(DateTime startTime, DateTime endTime)
        {
            TimeSpan ts = endTime - startTime;
            //string message = Convert.ToString(ts.TotalSeconds);
            string message = Convert.ToString(Math.Round(ts.TotalSeconds));
            return message;
        }

    }

    public enum DocumentStatus
    {
        Downloaded = 1,
        Preprocessed,
        Metadata,
        Extracted,
        Failed,
        Completed
    }

    public enum ErrorType
    {
        S = 1,
        B
    }
    public enum LogType
    {
        P = 1,
        E
    }
}
