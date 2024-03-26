using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReportGeneratorBlazorApp.Shared
{
    public class ReportFile
    {
        public required string SiteCollectionUrl { get; set; }
        public int ItemId{ get; set; }
        public Guid SpList { get; set; }        
        public int ReportItemId { get; set; }
        public Guid DocumentLib { get; set; }
        public ReportFileFormat FileFormat { get; set; }
    }
}
