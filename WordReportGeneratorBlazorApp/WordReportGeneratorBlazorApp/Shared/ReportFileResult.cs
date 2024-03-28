using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReportGeneratorBlazorApp.Shared
{
    public class ReportFileResult
    {        
        public required string CreatedFileName { get; set; }

        public required string FilePath { get; set; }
    }
}
