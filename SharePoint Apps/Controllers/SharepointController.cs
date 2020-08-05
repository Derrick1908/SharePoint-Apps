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
        [System.Web.Http.Route("api/getsharepointoken")]
        [HttpPost]
        public async Task<SharePointResponse> GetSharepointToken()
        {
            
            RequestModel requestModel = createRequests.CreateSharePointTokenRequestValues();
            return await createRequests.POSTAsync(requestModel);
        }

        [System.Web.Http.Route("api/CreateFolder")]
        [HttpPost]
        public async Task<SharePointResponse> CreateFolder(FolderModel folder)
        {
            SharePointResponse sharePointToken = await GetSharepointToken();
            RequestModel requestModel = createRequests.CreateSharePointFolderValues(folder);
            requestModel.token = sharePointToken.access_token;
            return await createRequests.POSTAsync(requestModel);
        }

    }
}
