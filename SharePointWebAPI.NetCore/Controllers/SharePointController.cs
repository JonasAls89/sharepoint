using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
//using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
using WebAPI.NetCore.Models;
using WebAPI.NetCore.Middleware;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WebAPI.NetCore.Controllers
{
    /// <summary>
    /// SharePoint Web API controller
    /// </summary>
    [Produces("application/json")]
    [Route("api/[controller]/[action]")]
    public class SharePointController : Controller
    {
        private string _username, _password;
        private readonly SharePointContext _context;
        private readonly ClientContext cc;



        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="config"></param>
        /// <param name="context"></param>
        public SharePointController(IConfiguration config, SharePointContext context)
        {
            string url = "";
            string _username = "";
            string _password = "";
            
            using (var file = System.IO.File.OpenText("helpers.json"))
            {
                var reader = new JsonTextReader(file);
                var jObject = JObject.Load(reader);
                url = jObject.GetValue("url").ToString();
                _username = jObject.GetValue("username").ToString();
                _password = jObject.GetValue("password").ToString();
            }
            
            _context = context;

            try
            {
                Console.WriteLine("Authenticating");
                this.cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password);
            }

            catch (NullReferenceException e)
            {
                System.Diagnostics.Debug.WriteLine("Exception occured whilst obtaining client context due to: " + e.Message);
                throw new ArgumentNullException(e.Message);
            }
        }
        /// <summary>
        /// Delete a site
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     DELETE /api/sharepoint/deletesite
        ///     {
        ///        "url": "https://....com/testteam1"
        ///     }
        /// </remarks>
        /// <param name="url">The site url to delete</param>
        /// <returns></returns>
        /// <response code="204">Returns success with No-content result</response>
        /// <response code="500">If the input parameter is null or empty</response>
        [HttpDelete("{url}")]
        [ProducesResponseType((int)HttpStatusCode.NotFound)]
        [ProducesResponseType((int)HttpStatusCode.RequestTimeout)]
        [ProducesResponseType((int)HttpStatusCode.NoContent)]
        [ProducesResponseType((int)HttpStatusCode.InternalServerError)]
        public async Task<IActionResult> DeleteSite(string url)
        {
            // Starting with ClientContext, the constructor requires a URL to the 
            // server running SharePoint.
            //string teamURL = @"https://dddevops.sharepoint.com/TeamSite1";
            try{    
                //Web web = context.Web;
                //context.Load(web);
                //context.Credentials = new NetworkCredential("khteh", "", "dddevops.onmicrosoft.com");
                Web web = cc.Web;
                // Retrieve the new web information. 
                cc.Load(web);
                //context.Load(newWeb);
                await cc.ExecuteQueryAsync();
                web.DeleteObject();
                await cc.ExecuteQueryAsync();
                return new NoContentResult();
            } catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, ex);
            }
        }
        /// <summary>
        /// Create a new site
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     POST /api/sharepoint/newsite
        ///     {
        ///        "param": {
        ///             "SiteCollectionURL": "https://....com",
        ///             "Title": "Test Team 1",
        ///             "URL": "TestTeam1",
        ///             "Description": "Test Team 1 description",
        ///             "Template": "STS#0"
        ///        }
        ///     }
        /// </remarks>
        /// <param name="param">Site creation parameters</param>
        /// <returns></returns>
        /// <response code="201">Returns success with the new site title</response>
        /// <response code="404">Returns resource not found if the ID of the new site is empty</response>
        /// <response code="500">If the input parameter is null or empty</response>
        [HttpPost]
        [ProducesResponseType((int)HttpStatusCode.NotFound)]
        [ProducesResponseType((int)HttpStatusCode.RequestTimeout)]
        [ProducesResponseType((int)HttpStatusCode.Created)]
        [ProducesResponseType((int)HttpStatusCode.InternalServerError)]
        public async Task<IActionResult> NewSite(SharePointParam param)
        {
            try{
            // Starting with ClientContext, the constructor requires a URL to the 
            // server running SharePoint.
            //string teamURL = @"https://dddevops.sharepoint.com";
            WebCreationInformation creation = new WebCreationInformation();
            //context.Credentials = new NetworkCredential("khteh", "", "dddevops.onmicrosoft.com");
            creation.Url = param.URL;
            creation.Title = param.Title;
            creation.Description = param.Description;
            creation.UseSamePermissionsAsParentSite = true;
            creation.WebTemplate = param.Template;//"STS#0";
            creation.Language = 1033;
            Web newWeb = cc.Web.Webs.Add(creation);
            // Retrieve the new web information. 
            cc.Load(newWeb, w => w.Id);
            //context.Load(newWeb);
            await cc.ExecuteQueryAsync();
            return StatusCode(newWeb.Id != Guid.Empty ? StatusCodes.Status201Created : StatusCodes.Status404NotFound);
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, ex);
            }
        }
#if false
        This section is commented out due to the fact that current Microsoft SharePoint Cliet component SDK does NOT support tennat administration for .Net Core
        [HttpGet]
        public List<SharePointItem> SiteCollections()
        {
            List<SharePointItem> results = new List<SharePointItem>();
            SPOSitePropertiesEnumerable prop = null;
            string tenantAdminURL = @"https://dddevops-admin.sharepoint.com/";
            using (ClientContext context = new ClientContext(tenantAdminURL))
            {
                SecureString securePassword = new SecureString();
                foreach (char c in password.ToCharArray())
                    securePassword.AppendChar(c);
                context.Credentials = new SharePointOnlineCredentials(username, securePassword);
                Tenant tenant = new Tenant(context);
                prop = tenant.GetSiteProperties(0, true);
                context.Load(prop);
                context.ExecuteQuery();
                foreach (SiteProperties sp in prop)
                {
                    SharePointItem item = new SharePointItem() { Title = sp.Title, URL = sp.Url };
                    results.Add(item);
                }
            }
            return results;
        }
#endif
        /// <summary>
        /// Retrieve all the sites of a specified site collection
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/sharepoint/sites
        ///     {
        ///        "url": "https://....com"
        ///     }
        /// </remarks>
        /// <!--<param name="url">Site collection URL</param>-->
        /// <returns>List of sites</returns>
        /// <response code="200">Returns success with the list of sites</response>
        /// <response code="500">Exception thrown in SharePoint server</response>
        [HttpGet]
        [ProducesResponseType(typeof(List<SharePointParam>), (int)HttpStatusCode.OK)]
        [ProducesResponseType((int)HttpStatusCode.InternalServerError)]
        public async Task<IActionResult> Sites()
        {
            List<SharePointParam> results = new List<SharePointParam>();
            try
            {
                // Root Web Site 
                Web spRootWebSite = cc.Web;
                // Collecction of Sites under the Root Web Site 
                WebCollection spSites = spRootWebSite.Webs;
                // Loading operations         
                cc.Load(spRootWebSite);
                cc.Load(spSites);
                await cc.ExecuteQueryAsync();
                List<Task> tasks = new List<Task>();
                // We need to iterate through the $spoSites Object in order to get individual sites information 
                foreach (Web site in spSites)
                {
                    Console.WriteLine("Writing sites " + site);
                    cc.Load(site);
                    tasks.Add(cc.ExecuteQueryAsync());
                }
                Task.WaitAll(tasks.ToArray());
                foreach (Web site in spSites)
                {
                    SharePointParam item = new SharePointParam() { Title = site.Title, URL = site.Url };
                    results.Add(item);
                }
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, ex);
            }
            return new OkObjectResult(results);
        }
        /// <summary>
        /// Retrieve the available templates of a site collection
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/sharepoint/templates
        ///     {
        ///        "url": "https://....com"
        ///     }
        /// </remarks>
        /// <param name="url">URL of a site collection to retrieve the templates from</param>
        /// <returns>List of template information</returns>
        /// <response code="200">Returns success with the list of template information</response>
        /// <response code="500">Exception thrown in SharePoint server</response>
        [HttpGet]
        [ProducesResponseType(typeof(List<SharePointTemplate>), (int)HttpStatusCode.OK)]
        [ProducesResponseType((int)HttpStatusCode.InternalServerError)]
        public async Task<IActionResult> Templates(string url)
        {
            List<SharePointTemplate> results = new List<SharePointTemplate>();
            try
            {
                Web web = cc.Web;
                // LCID: https://msdn.microsoft.com/en-us/library/ms912047%28v=winembedded.10%29.aspx?f=255&MSPPError=-2147217396
                WebTemplateCollection templates = web.GetAvailableWebTemplates(1033, false);
                cc.Load(templates);
                //Execute the query to the server    
                await cc.ExecuteQueryAsync();
                // Loop through all the list templates    
                foreach (WebTemplate template in templates)
                {
                    SharePointTemplate item = new SharePointTemplate() { ID = template.Id, Title = template.Title, Name = template.Name, Description = template.Description };
                    results.Add(item);
                }   
            } catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, ex);
            }
            return new OkObjectResult(results);
        }
    }
}