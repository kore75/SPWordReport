using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.Resource;
using PnP.Core.Services.Builder.Configuration;
using PnP.Core.Services;
using WordReportGeneratorBlazorApp.Shared;
using Microsoft.Extensions.Options;
using PnP.Core.Auth;
using System.Collections.Generic;

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
            using (var context = await createSiteContextForUser())
            {
                var list=await context.Web.Lists.GetByTitleAsync(listName);
                foreach(var item in list.Items.ToList()) 
                { 
                    itemDatas.Add(new ListItemData { Id=item.Id,Title=item.Title });
                }                
                //Dictionary<string, object> values = new Dictionary<string, object>
                //{
                //    { "Title", "PnP Rocks!" }
                //};

                //await list.Items.AddBatchAsync(values);
                //await list.Items.AddBatchAsync(values);
                //await list.Items.AddBatchAsync(values);

                //// Send batch to the server
                //await context.ExecuteAsync();
            }

            

            return itemDatas;
        }

        //// GET api/<ReportGeneratorController>/5
        //[HttpGet("{id}")]
        //public string Get(int id)
        //{
        //    return "value";
        //}

        // POST api/<ReportGeneratorController>
        [HttpPost]
        public void Post([FromBody] ReportFile reportFile)
        {

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
