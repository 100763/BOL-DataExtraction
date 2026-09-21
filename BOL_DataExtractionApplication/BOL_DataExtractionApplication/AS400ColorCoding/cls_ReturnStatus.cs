using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using BOLDocumentDataExtraction.CommonModule;

namespace BOLDocumentDataExtraction.AS400ColorCoding
{
  public  class cls_ReturnStatus
    {
        public string Message { get; set; }
        public string Status { get; set; }

        public string Value { get; set; }

        public bool IsStatus { get; set; }
        public int BlankLineItemNumber { get; set; }
    }
}
