using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOLDocumentDataExtraction
{
    public class ServiceResponse
    {
        public ServiceResponseStatus Status
        {
            get; set;
        }
        public string Message
        {
            get; set;
        }
        public DataTable Response
        {
            get; set;
        }
    }
    public enum ServiceResponseStatus
    {
        Success = 1,
        Failed
    }
}
