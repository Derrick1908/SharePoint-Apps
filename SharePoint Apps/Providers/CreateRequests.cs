using Newtonsoft.Json;
using SharePoint_Apps.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SharePoint_Apps.Providers
{
    public class CreateRequests
    {
        /// <summary>
        /// Description : This method creates the URL for the POST Request that will hit the
        ///               SharePoint API inorder to retrieve the OAuth Token that will be used in Future Requests to Create Folders/Upload Files
        ///               along with the needed Body Content in the form of form-url-encoded
        /// </summary>
        /// <returns> The URL for the POST Request and the needed body values</returns>
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
                    Values = values,
                    type = 3
                };

                return requestModel;
            }
            catch(Exception)
            {
                return null;
            }
        }


        /// <summary>
        /// Description : This method creates the URL for the POST Request that will hit the
        ///               SharePoint API inorder to create a Folder with the Supplied Name in the Shared Documents Folder.
        ///               along with the needed Body Content in the form of application/json;odata=verbose
        /// </summary>
        /// <returns> The URL for the POST Request and the needed body values</returns>
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
                    body = myJson,
                    type = 1
                };

                return requestModel;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Description : This method creates the URL for the POST Request that will hit the
        ///               SharePoint API inorder to get the Form Digest Value.
        /// </summary>
        /// <returns> The URL for the POST Request </returns>
        public RequestModel CreateSharePointFormDigestRequestValues()
        {
            try
            {
                var parentURL = ConfigurationManager.AppSettings["Parent SharePoint URL"].ToString();
                var subURL_1 = ConfigurationManager.AppSettings["Sub Part 1 URL"].ToString();
                var FormDigest_subURL_2 = ConfigurationManager.AppSettings["Form Digest Part 2 URL"].ToString();


                string formDigestURL = parentURL + subURL_1 + FormDigest_subURL_2;
                RequestModel requestModel = new RequestModel
                {
                    URL = formDigestURL,
                    type = 2
                };

                return requestModel;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Description : A Commmon Method that creates POST Requests based on Different
        ///               Need that will be called with the appropriate details
        /// </summary>
        /// <param name="credentials"></param>
        /// <returns>The appropriate result which could range from token value/form digest value</returns>
        public async Task<object> POSTAsync(RequestModel credentials)
        {
            using (HttpClient client = new HttpClient())
            {
                HttpContent content;
                if (credentials.body != null && credentials.type == 1) //Type 1 indicates that the POST Request will be for Uploading a Folder
                {
                    //This Section Creates a POST Request for Creating a Folder

                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", credentials.token);
                    MediaTypeWithQualityHeaderValue acceptHeader = MediaTypeWithQualityHeaderValue.Parse("application/json;odata=verbose;charset=utf-8");
                    client.DefaultRequestHeaders.Accept.Add(acceptHeader);
                    content = new StringContent(credentials.body);
                    content.Headers.ContentType = acceptHeader;                    
                    var result = await client.PostAsync(credentials.URL, content);
                    return result;
                }
                else if(credentials.type == 2)  //Type 2 indicates that the POST Request will be for Form Digest Value
                {
                    //This Section Creates a POST Request for Getting the Form Digest Value

                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", credentials.token);
                    MediaTypeWithQualityHeaderValue acceptHeader = MediaTypeWithQualityHeaderValue.Parse("application/json;odata=verbose");
                    client.DefaultRequestHeaders.Accept.Add(acceptHeader);
                    content = null;
                    var result = await client.PostAsync(credentials.URL, content);
                    string resultContent = await result.Content.ReadAsStringAsync();                    
                    return resultContent;
                }
                else            //Type 3 indicates that the POST Request will be for Getting the Token
                {
                    //This Section Creates a POST Request for Retrieving a Token

                    content = new FormUrlEncodedContent(credentials.Values);
                    var result = await client.PostAsync(credentials.URL, content);
                    string resultContent = await result.Content.ReadAsStringAsync();
                    SharePointResponse tokenResponse = JsonConvert.DeserializeObject<SharePointResponse>(resultContent);
                    return tokenResponse;
                }                               
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