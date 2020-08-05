using Newtonsoft.Json;
using SharePoint_Apps.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SharePoint_Apps.Providers
{
    public class CreateRequests
    {
        public RequestModel CreateSharePointTokenRequestValues()
        {
            try
            {
                var tokenURL = ConfigurationManager.AppSettings["Token URL"].ToString();
                var clientId = ConfigurationManager.AppSettings["Client Id"].ToString();
                var clientSecret = ConfigurationManager.AppSettings["Client Secret"].ToString();
                var tenantId = ConfigurationManager.AppSettings["Tenant Id"].ToString();
                var tenantName = ConfigurationManager.AppSettings["Tenant Name"].ToString();
                var resource = ConfigurationManager.AppSettings["resource"].ToString();
                var client_id = ConfigurationManager.AppSettings["client_id"].ToString();

                resource = resource.Replace("*tenantname*", tenantName).Replace("*tenantid*", tenantId);
                client_id = client_id.Replace("*clientid*", clientId).Replace("*tenantid*", tenantId);
                var values = new Dictionary<string, string>{
                {"grant_type", "client_credentials"},
                {"resource", resource},
                {"client_id", client_id},
                {"client_secret", clientSecret},
            };
                tokenURL = tokenURL.Replace("*tenantid*", tenantId);
                RequestModel requestModel = new RequestModel
                {
                    URL = tokenURL,
                    Values = values
                };

                return requestModel;
            }
            catch(Exception)
            {
                return null;
            }
        }
        
        public RequestModel CreateSharePointFolderValues(FolderModel folders)
        {
            try
            {
                var parentURL = ConfigurationManager.AppSettings["Parent SharePoint URL"].ToString();
                var subURL_1 = ConfigurationManager.AppSettings["Sub Part 1 URL"].ToString();
                var subURL_2 = ConfigurationManager.AppSettings["Sub Part 2 URL"].ToString();
                var folder = ConfigurationManager.AppSettings["Folder Directory URL"].ToString();


                string folderURL = parentURL + subURL_1 + subURL_2 + "folders";
                string myJson = "{'__metadata': {'type': 'SP.Folder'},'ServerRelativeUrl': '" + subURL_1 + folder + folders.FolderName + "'}";

                RequestModel requestModel = new RequestModel
                {
                    URL = folderURL,
                    body = myJson
                };

                return requestModel;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public async Task<SharePointResponse> POSTAsync(RequestModel credentials)
        {
            using (HttpClient client = new HttpClient())
            {
                HttpContent content;
                if (credentials.body != null)
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", credentials.token);
                    MediaTypeWithQualityHeaderValue acceptHeader = MediaTypeWithQualityHeaderValue.Parse("application/json;odata=verbose");
                    client.DefaultRequestHeaders.Accept.Add(acceptHeader);
                    content = new StringContent(credentials.body);
                }
                else
                {
                    content = new FormUrlEncodedContent(credentials.Values);
                }
                var result = await client.PostAsync(credentials.URL, content);
                string resultContent = await result.Content.ReadAsStringAsync();
                SharePointResponse tokenResponse = JsonConvert.DeserializeObject<SharePointResponse>(resultContent);
                return tokenResponse;
            }
        }
        public async Task<object> GETAsync(RequestModel credentials)
        {

            var client = new HttpClient();            
            var result = await client.GetAsync(credentials.URL);
            string resultContent = await result.Content.ReadAsStringAsync();
            return resultContent;
        }
    }
}