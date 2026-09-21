using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOLDocumentDataExtraction.CommonModule
{
    public class ErrorLog
    {
        public long DocumentID
        {
            get; set;
        }
        public ErrorType ErrorType
        {
            get; set;
        }
        public string ErrorDetails
        {
            get; set;
        }
        public string ModuleName
        {
            get; set;
        }
        public string FunctionName
        {
            get; set;
        }


    }
}
