using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOLDocumentDataExtraction.CommonModule
{
    public class DocumentMaster
    {
        public long DocumentID
        {
            get; set;
        }
        public string ServerName
        {
            get; set;
        }
        public string ServerIP
        {
            get; set;
        }
        public string SourcePath
        {
            get; set;
        }
        public string DestinationPath
        {
            get; set;
        }
        public string FileName
        {
            get; set;
        }
        public int TemplateNumber
        {
            get; set;
        }
        public int CurrentStatus
        {
            get; set;
        }
        public DateTime HeaderFieldSearch
        {
            get; set;
        }
        public DateTime LineFieldSearch
        {
            get; set;
        }
        public DateTime Verify
        {
            get; set;
        }
        public string Verify_User_Name
        {
            get; set;
        }
        public bool Verify_Locked
        {
            get; set;
        }
        public DateTime RPA
        {
            get; set;
        }
    }
}
