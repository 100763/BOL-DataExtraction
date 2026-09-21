using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BOLDocumentDataExtraction.CommonModule;

namespace BOLDocumentDataExtraction.AS400ColorCoding
{
   public class cls_HeaderField
    {
        public string SrNo { get; set; }
        public string FieldName
        { get; set; }

        public string FieldValue
        { get; set; }

        public string FieldColor { get; set; }

        public double RowNo { get; set; }

    }
}
