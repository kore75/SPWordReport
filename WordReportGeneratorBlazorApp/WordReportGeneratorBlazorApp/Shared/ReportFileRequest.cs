using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReportGeneratorBlazorApp.Shared
{
    public class ReportFileRequest
    {        
        public int ItemId{ get; set; }
        public Guid SpListGuid { get; set; }        
        public int ReportItemId { get; set; }
        public Guid DocumentLibGuid { get; set; }
    }
}
