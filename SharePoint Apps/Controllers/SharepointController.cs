using SharePoint_Apps.Models;
using SharePoint_Apps.Providers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

namespace SharePoint_Apps.Controllers
{
    public class SharepointController : ApiController
    {
        CreateRequests createRequests = new CreateRequests();
        
        /// <summary>
        /// Description : POST Request to Get the Token based on Successful Credentials supplied
        ///               like the Client ID and Client Secret
        /// </summary>
        /// <returns>the Token on supplying correct credentials</returns>
        [System.Web.Http.Route("api/getsharepointoken")]
        [HttpPost]
        public async Task<SharePointResponse> GetSharepointToken()
        {
            
            RequestModel requestModel = createRequests.CreateSharePointTokenRequestValues();
            return (SharePointResponse) await createRequests.POSTAsync(requestModel);
        }

        /// <summary>
        /// Description : POST Requests to Create a Folder on Sharepoint. Internally calls
        ///               the Get Token Function.
        /// </summary>
        /// <param name="folder"></param>
        /// <returns></returns>
        [System.Web.Http.Route("api/CreateFolder")]
        [HttpPost]        
        public async Task<object> CreateFolder(FolderModel folder)
        {
            var configuration = new HttpConfiguration();
            var request = new HttpRequestMessage();
            request.SetConfiguration(configuration);
            try
            {
                SharePointResponse sharePointToken = await GetSharepointToken();
                RequestModel requestModel = createRequests.CreateSharePointFolderValues(folder);
                requestModel.token = sharePointToken.access_token;
                await createRequests.POSTAsync(requestModel);
                return request.CreateResponse(HttpStatusCode.OK, "Successfully Created Folder");
            }
            catch(Exception)
            {
                return new HttpResponseMessage(HttpStatusCode.InternalServerError);
            }
        }

        /// <summary>
        /// Description : Gets the Form Digest Value using a Valid Token supplied along 
        ///               with the Request
        /// </summary>
        /// <returns>Form Digest Value</returns>
        [System.Web.Http.Route("api/GetFormDigest")]
        [HttpPost]
        public async Task<SharePointResponse> GetFormDigest ()
        {
            var configuration = new HttpConfiguration();
            var request = new HttpRequestMessage();
            request.SetConfiguration(configuration);
            try
            {
                SharePointResponse sharePointToken = await GetSharepointToken();                
                RequestModel requestModel = createRequests.CreateSharePointFormDigestRequestValues();
                requestModel.token = sharePointToken.access_token;
                string formDigestResult = (string) await createRequests.POSTAsync(requestModel);
                string[] formDigest = formDigestResult.Split('\"');
                for (var i = 0; i < formDigest.Length; i++)
                {
                    if (formDigest[i].Equals("FormDigestTimeoutSeconds"))
                    {
                        sharePointToken.formDigestTimeout = int.Parse(formDigest[i + 1].Replace(":", "").Replace(",",""));                        
                    }
                    if (formDigest[i].Equals("FormDigestValue"))
                    {
                        sharePointToken.formDigestValue = formDigest[i + 2];
                        break;
                    }
                }
                return sharePointToken;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
}
