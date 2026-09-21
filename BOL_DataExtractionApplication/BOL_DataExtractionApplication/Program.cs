using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOLDocumentDataExtraction
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
            try
            {
                Process process = Process.GetCurrentProcess();
                DataExtraction d = new DataExtraction();
                d.StaticVariable_Initialize();
                d.StartDataExtractionService();
                d.StaticVariable_Release();
                d = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                process.Kill();
            }
            catch (Exception Ex)
            {
                if (EventLog.SourceExists("BOLDataExtractionServiceSource"))
                {
                    EventLog.WriteEntry("BOLDataExtractionServiceSource", Ex.Message, EventLogEntryType.Error, 103);
                }
            }

        }
    }
}
