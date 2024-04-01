using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.Resource;
using PnP.Core.Auth;
using PnP.Core.Model;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.Net.Mime;
using WordReportGeneratorBlazorApp.Shared;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace WordReportGeneratorBlazorApp.Server.Controllers
{
    [Authorize]
    [ApiController]
    [Route("[controller]")]
    [RequiredScope(RequiredScopesConfigurationKey = "AzureAd:Scopes")]
    public class ReportGeneratorController : ControllerBase
    {
        private readonly ILogger<ReportGeneratorController> _logger;
        private readonly IPnPContextFactory _pnpContextFactory;
        private readonly ITokenAcquisition _tokenAcquisition;
        private readonly PnPCoreOptions _pnpCoreOptions;

        public ReportGeneratorController(IPnPContextFactory pnpContextFactory,
            ILogger<ReportGeneratorController> logger,
            ITokenAcquisition tokenAcquisition,
            IOptions<PnPCoreOptions> pnpCoreOptions)
        {
            _pnpContextFactory = pnpContextFactory;
            _logger = logger;
            _tokenAcquisition = tokenAcquisition;
            _pnpCoreOptions = pnpCoreOptions.Value;
        }

        private async Task<PnPContext> createSiteContextForUser()
        {
            var siteUrl = new Uri(_pnpCoreOptions.Sites["ReportSite"].SiteUrl);

            return await _pnpContextFactory.CreateAsync(siteUrl,
                            new ExternalAuthenticationProvider((resourceUri, scopes) =>
                            {
                                return _tokenAcquisition.GetAccessTokenForUserAsync(scopes,user:this.User);
                            }
                            ));
        }

        private async Task<PnPContext> createSiteContextApp()
        {
            var siteUrl = new Uri(_pnpCoreOptions.Sites["ReportSite"].SiteUrl);

            return await _pnpContextFactory.CreateAsync(siteUrl,
                            new ExternalAuthenticationProvider((resourceUri, scopes) =>
                            {
                                return _tokenAcquisition.GetAccessTokenForAppAsync(scopes[0]);
                            }
                            ));
        }

        //// GET: api/<ReportGeneratorController>
        [HttpGet]
        public async Task<IEnumerable<ListItemData>> Get([FromQuery]string listName)
        {
            List<ListItemData> itemDatas=new List<ListItemData>();
            bool createSampleData = false;
            using (var context = await createSiteContextForUser())
            {
                var list=await context.Web.Lists.GetByTitleAsync(listName);
                foreach(var item in list.Items.ToList()) 
                { 
                    itemDatas.Add(new ListItemData { Id=item.Id,Title=item.Title });
                }
                if(createSampleData) 
                {
                    Dictionary<string, object> values = new Dictionary<string, object>
                    {
                        { "Title", "PnP Rocks!" }
                    };

                    await list.Items.AddBatchAsync(values);
                    await list.Items.AddBatchAsync(values);
                    await list.Items.AddBatchAsync(values);

                    // Send batch to the server
                    await context.ExecuteAsync();
                }

            }

            

            return itemDatas;
        }

        // POST api/<ReportGeneratorController>
        [HttpPost]
        [Consumes(MediaTypeNames.Application.Json)]
        [ProducesResponseType(StatusCodes.Status201Created)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<ActionResult<ReportFileResult>> Post([FromBody] ReportFileRequest reportFile)
        {
            try
            {
                using (var context = await createSiteContextForUser())
                {                  
                    var spList = await context.Web.Lists.GetByIdAsync(reportFile.SpListGuid);
                    if (spList == null)
                    {
                        return NotFound($"Source list with id:{reportFile.SpListGuid} not found");
                    }
                    var item = await spList.Items.QueryProperties(x => x.Values["Author"]).FirstOrDefaultAsync(x => x.Id == reportFile.ItemId);
                    if (item == null)
                    {
                        return NotFound($"Source list item with id:{reportFile.ItemId} not found");
                    }
                    //var destDocLib = await context.Web.GetFolderByIdAsync(reportFile.DocumentLibGuid);
                    var destDocLib = await context.Web.Lists.GetByIdAsync(reportFile.DocumentLibGuid);
                    if (destDocLib == null)
                    {
                        return NotFound($"Template DocLib with id:{reportFile.DocumentLibGuid} not found");
                    }
                    var reportTemplate = await destDocLib.Items.GetByIdAsync(1, li => li.All, li => li.File);
                    if (reportTemplate == null)
                    {
                        return NotFound($"Template File with id:{reportFile.ReportItemId} not found");
                    }
                    string templateFileExtennsion = Path.GetExtension(reportTemplate.File.Name);
                    if (string.Compare(templateFileExtennsion, ".docx", true) != 0)
                    {
                        return BadRequest("Template in wrong format, has to be docx");
                    }
                    // Create a new document
                    WordDocument document = new WordDocument(reportTemplate.File.GetContent(), FormatType.Docx);
                    string[] fieldNames = new string[item.Values.Count];
                    string[] fieldValues= new string[item.Values.Count];
                    int idx = 0;
                    foreach (var itemValue in item.Values)
                    {
                        fieldNames[idx] = itemValue.Key;
                        fieldValues[idx] = Convert.ToString(itemValue.Value) ?? "";
                        idx++;
                    }                                   
                    //Performs the mail merge
                    document.MailMerge.Execute(fieldNames, fieldValues);
                    //
                    DocIORenderer render = new DocIORenderer();
                    //Sets Chart rendering Options.
                    render.Settings.ChartRenderingOptions.ImageFormat = Syncfusion.OfficeChart.ExportImageFormat.Jpeg;
                    //create Pdf
                    PdfDocument pdfDocument = render.ConvertToPDF(document);
                    //Saves the Word document to MemoryStream
                    using (MemoryStream stream = new MemoryStream())
                    {
                        pdfDocument.Save(stream);
                        //Closes the Word document
                        document.Close();
                        stream.Position = 0;
                        var baseUri = context.Web.Url;
                        string baseUrl = $"{baseUri.Scheme}://{baseUri.DnsSafeHost}";
                       
                        var newAttachment = await item.AttachmentFiles.AddAsync($"Report_{item.Id}_{DateTime.Now:yyyyMMddHHmmss}.pdf", stream);
                        string createdUri = $"{baseUrl}{newAttachment.ServerRelativeUrl}";
                        ReportFileResult newReportFile = new ReportFileResult() { CreatedFileName = newAttachment.FileName, FilePath = createdUri };                                                              
                        return Created(createdUri, newReportFile);
                    }                    
                }
            }
            catch (Exception exception)
            {
                return BadRequest(exception.Message);
            }            
        }

        //// PUT api/<ReportGeneratorController>/5
        //[HttpPut("{id}")]
        //public void Put(int id, [FromBody] string value)
        //{
        //}

        //// DELETE api/<ReportGeneratorController>/5
        //[HttpDelete("{id}")]
        //public void Delete(int id)
        //{
        //}
    }
}
