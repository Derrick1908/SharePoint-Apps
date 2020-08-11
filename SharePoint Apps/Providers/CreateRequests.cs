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
using System.Web.Compilation;

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
        /// TO DO: To update Sub Folder Names later based on what is finally needed.
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

                folders.SubFolders = new List<string>();
                if (folders.type == 1)   //Type 1 indicates Clients
                {
                    folders.SubFolders.Add("FF&E");
                    folders.SubFolders.Add("OS&E");
                    folders.SubFolders.Add("Logos");
                    folders.SubFolders.Add("Profile Pics");
                }
                else          //Type 2 Indicates Vendors (Temp. To be Updated)
                {
                    folders.SubFolders.Add("FF&E");

                }
                List<string> myJson_2 = new List<string>();
                for(int i=0; i<folders.SubFolders.Count;i++)
                {
                    myJson_2.Add("{'__metadata': {'type': 'SP.Folder'},'ServerRelativeUrl': '" + subURL_1 + folder + folders.FolderName + "/" + folders.SubFolders[i] + "'}");
                }
                RequestModel requestModel = new RequestModel
                {
                    URL = folderURL,
                    body = myJson,
                    body_2 = myJson_2,
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
        ///               SharePoint API inorder to Delete a Folder with the Supplied Name in the Shared Documents Folder.
        ///               along with the needed details.
        /// </summary>
        /// <returns> The URL for the POST Request with the needed values</returns>
        public RequestModel DeleteSharePointFolderValues(FolderModel folders)
        {
            try
            {
                var parentURL = ConfigurationManager.AppSettings["Parent SharePoint URL"].ToString();
                var subURL_1 = ConfigurationManager.AppSettings["Sub Part 1 URL"].ToString();
                var subURL_2 = ConfigurationManager.AppSettings["Sub Part 2 URL"].ToString();
                var subfolder = ConfigurationManager.AppSettings["Folder Directory URL"].ToString();
                var folderURL = ConfigurationManager.AppSettings["Folder Relative URL"].ToString();

                string folderRelativeURL = string.Format(folderURL,subURL_1 + subfolder + folders.FolderName);
                string folderDeleteURL = parentURL + subURL_1 + subURL_2 + folderRelativeURL;

                RequestModel requestModel = new RequestModel
                {
                    URL = folderDeleteURL,
                    type = 4
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
        ///               SharePoint API inorder to Upload a File to a particular Folder with the Supplied Name in the Shared Documents Folder.
        ///               along with the needed details.
        /// </summary>
        /// <returns> The URL for the POST Request with the needed values</returns>
        public RequestModel UploadFileSharePointValues(FolderModel folders)
        {
            try
            {
                var parentURL = ConfigurationManager.AppSettings["Parent SharePoint URL"].ToString();
                var subURL_1 = ConfigurationManager.AppSettings["Sub Part 1 URL"].ToString();
                var subURL_2 = ConfigurationManager.AppSettings["Sub Part 2 URL"].ToString();
                var subfolder = ConfigurationManager.AppSettings["Folder Directory URL"].ToString();
                var folderURL = ConfigurationManager.AppSettings["Folder Relative URL"].ToString();

                string fileTempURL = subURL_1 + subfolder + folders.FolderName + "/";
                if (folders.path != null)   //Incase the Document is within a particular Folder under that particular Client/Vendor etc.
                    fileTempURL += folders.path; //Those Sub Folders will also be added to the Path.

                var fileURL = ConfigurationManager.AppSettings["File Relative URL"].ToString();
                string folderRelativeURL = string.Format(folderURL, fileTempURL);
                string fileuploadURL = parentURL + subURL_1 + subURL_2 + folderRelativeURL;
                fileuploadURL += string.Format(fileURL, folders.fileName);
                RequestModel requestModel = new RequestModel
                {
                    URL = fileuploadURL,
                    type = 5
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
        ///               SharePoint API inorder to Delete a File from a particular Folder with the Supplied Name in the Shared Documents Folder
        ///               along with the needed details.
        /// </summary>
        /// <returns> The URL for the POST Request with the needed values</returns>
        public RequestModel DeleteFileSharePointValues(FolderModel folders)
        {
            try
            {
                var parentURL = ConfigurationManager.AppSettings["Parent SharePoint URL"].ToString();
                var subURL_1 = ConfigurationManager.AppSettings["Sub Part 1 URL"].ToString();
                var subURL_2 = ConfigurationManager.AppSettings["Sub Part 2 URL"].ToString();
                var subfolder = ConfigurationManager.AppSettings["Folder Directory URL"].ToString();
                var fileURL = ConfigurationManager.AppSettings["Get File Relative URL"].ToString();

                string fileTempURL = subURL_1 + subfolder + folders.FolderName + "/";
                if (folders.SubFolders != null)   //Incase the Document is within a particular Folder under that particular Client/Vendor etc.
                    for (int i = 0; i < folders.SubFolders.Count; i++)
                        fileTempURL += "/" + folders.SubFolders[i]; //Those Sub Folders will also be added to the Path.
                fileTempURL += "/" + folders.fileName;
                string fileRelativeURL = string.Format(fileURL, fileTempURL);
                string fileDeleteURL = parentURL + subURL_1 + subURL_2 + fileRelativeURL;

                RequestModel requestModel = new RequestModel
                {
                    URL = fileDeleteURL,
                    type = 7
                };

                return requestModel;
            }
            catch (Exception)
            {
                return null;
            }
        }
        /// <summary>
        /// Description : This method creates the URL for the GET Request that will hit the
        ///               SharePoint API inorder to List a Folder (name supplied) Contents in the Shared Documents Folder.        
        /// </summary>
        /// <returns> The URL for the GET Request with the needed values</returns>
        public RequestModel GetFolderContentSharePointValues(FolderModel folders)
        {
            try
            {
                RequestModel requestModel = DeleteSharePointFolderValues(folders);
                requestModel.URL2 = requestModel.URL + "/Folders";
                requestModel.URL += "/Files";                
                requestModel.type = 6;
                return requestModel;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>        
        /// Description : This method creates the URL for the GET Request that will hit the
        ///               SharePoint API inorder to Check if the Folder (name supplied) exists in the Shared Documents Folder.
        ///               Note that it Calls DeleteFolder Method to build the URL and due to this it can also
        ///               receive sub folders names as parameters in the call.
        /// </summary>
        /// <returns> The URL for the GET Request with the needed values</returns>
        public RequestModel FolderExistsSharePointValues(FolderModel folders)
        {
            try
            {
                RequestModel requestModel = DeleteSharePointFolderValues(folders);
                requestModel.URL += "/ListItemAllFields";
                requestModel.type = 8;
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
                if (credentials.body != null && credentials.type == 1) //Type 1 indicates that the POST Request will be for Creating a Folder
                {
                    //This Section Creates a POST Request for Creating a Folder on Shared Documents

                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", credentials.token);
                    MediaTypeWithQualityHeaderValue acceptHeader = MediaTypeWithQualityHeaderValue.Parse("application/json;odata=verbose;charset=utf-8");
                    client.DefaultRequestHeaders.Accept.Add(acceptHeader);
                    content = new StringContent(credentials.body);
                    content.Headers.ContentType = acceptHeader;                    
                    var result = await client.PostAsync(credentials.URL, content);
                    if(credentials.body_2!=null)
                    {
                        for(var i=0;i<credentials.body_2.Count;i++)
                        {
                            content = new StringContent(credentials.body_2[i]);
                            content.Headers.ContentType = acceptHeader;
                            await client.PostAsync(credentials.URL, content);
                        }
                    }
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
                else if(credentials.type == 3)  //Type 3 indicates that the POST Request will be for Getting the Token
                {
                    //This Section Creates a POST Request for Retrieving a Token

                    content = new FormUrlEncodedContent(credentials.Values);
                    var result = await client.PostAsync(credentials.URL, content);
                    string resultContent = await result.Content.ReadAsStringAsync();
                    SharePointResponse tokenResponse = JsonConvert.DeserializeObject<SharePointResponse>(resultContent);
                    return tokenResponse;
                }
                else if (credentials.type == 4)        //Type 4 indicates that the POST Request will be for Deleting a Folder on Shared Documents
                {
                    //This Section Creates a POST Request for Deleting a Folder

                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", credentials.token);
                    client.DefaultRequestHeaders.Add("X-RequestDigest", credentials.formDigestValue);
                    client.DefaultRequestHeaders.Add("X-HTTP-Method", "DELETE");
                    content = null;
                    var result = await client.PostAsync(credentials.URL, content);
                    return result;
                }
                else if (credentials.type == 5)        //Type 5 indicates that the POST Request will be for Uploading a File to a Particular Folder on Shared Documents
                {
                    //This Section Creates a POST Request for Uploading a File to a Particular Folder
                    MultipartFormDataContent form = new MultipartFormDataContent();
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", credentials.token);
                    client.DefaultRequestHeaders.Add("X-RequestDigest", credentials.formDigestValue);
                    form.Add(credentials.httpPostedFile);
                    var result = await client.PostAsync(credentials.URL, form);
                    return result;
                }

                else                               //Type 7 indicates that the POST Request will be for Deleting a File from a Particular Folder on Shared Documents
                {
                    //This Section Creates a POST Request for Deleting a File from a Particular Folder
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", credentials.token);
                    client.DefaultRequestHeaders.Add("X-RequestDigest", credentials.formDigestValue);
                    client.DefaultRequestHeaders.Add("X-HTTP-Method", "DELETE");
                    content = null;
                    var result = await client.PostAsync(credentials.URL, content);
                    return result;
                }
            }
        }

        /// <summary>
        /// Description : A Commmon Method that creates GET Requests based on Different
        ///               Need that will be called with the appropriate details
        /// </summary>
        /// <param name="credentials"></param>
        /// <returns>The appropriate result which could range from getting folder contents or checking if folder is empty</returns>
        public async Task<object> GETAsync(RequestModel credentials)
        {
            using (HttpClient client = new HttpClient())
            {
                if (credentials.type == 6)          //Type 6 indicates that the GET Request will be for Retrieving the contents of a Particular Folder exist on Shared Documents
                {
                    //This Section Creates a GET Request for Retrieving the Contents of a Particular Folder
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", credentials.token);
                    MediaTypeWithQualityHeaderValue acceptHeader = MediaTypeWithQualityHeaderValue.Parse("application/json;odata=verbose");
                    client.DefaultRequestHeaders.Accept.Add(acceptHeader);
                    var result = await client.GetAsync(credentials.URL);    //Files Result
                    var result2 = await client.GetAsync(credentials.URL2);  //Folders Result
                    
                    string resultContent = await result.Content.ReadAsStringAsync();
                    string resultContent2 = await result2.Content.ReadAsStringAsync();
                    return resultContent + "@" + resultContent2;  //Acts as a separator between both the Results
                }
                else   //Type 8 indicates that the GET Request will be for Checking if a a Particular Folder exist on Shared Documents or not
                {
                    //This Section Creates a GET Request for checking if a Particular Folder exists on Sharepoint Shared Documents or not.
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", credentials.token);
                    MediaTypeWithQualityHeaderValue acceptHeader = MediaTypeWithQualityHeaderValue.Parse("application/json;odata=verbose");
                    client.DefaultRequestHeaders.Accept.Add(acceptHeader);
                    var result = await client.GetAsync(credentials.URL);
                    string resultContent = await result.Content.ReadAsStringAsync();
                    return resultContent;
                }
            }
        }
    }
}