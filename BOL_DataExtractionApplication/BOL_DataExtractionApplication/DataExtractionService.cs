using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BOLDocumentDataExtraction
{
    partial class DataExtractionService : ServiceBase
    {
        Thread Worker;
        AutoResetEvent StopRequest = new AutoResetEvent(false);
        public DataExtractionService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            // TODO: Add code here to start your service.
            // Start the worker thread
            Worker = new Thread(DoWork);
            Worker.SetApartmentState(ApartmentState.STA);
            Worker.Start();
        }

        protected override void OnStop()
        {
            // TODO: Add code here to perform any tear-down necessary to stop your service.
            StopRequest.Set();
            Worker.Join();
        }
        private void DoWork(object arg)
        {
            // Worker thread loop
            for (; ; )
            {
                // Run this code once every 10 seconds or stop right away if the service 
                // is stopped
                if (StopRequest.WaitOne(10000)) return;
                DataExtraction d = new DataExtraction();
                d.StaticVariable_Initialize();
                d.StartDataExtractionService();
                d.StaticVariable_Release();
                d = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                // Do work...
                //...
            }
        }
    }
}
