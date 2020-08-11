using SharePoint_Apps.Models;
using SharePoint_Apps.Providers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
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
            try
            {
                RequestModel requestModel = createRequests.CreateSharePointTokenRequestValues();
                if (requestModel != null)
                    return (SharePointResponse)await createRequests.POSTAsync(requestModel);
                else
                    throw new Exception();
            }
            catch(Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Description : POST Requests to Create a Folder on Sharepoint. Internally calls
        ///               the Get Token Function.
        /// </summary>
        /// <param name="folder"></param>
        /// <returns>Whether the Folder is created successfully or not.</returns>
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
                if (sharePointToken != null)
                {
                    RequestModel requestModel = createRequests.CreateSharePointFolderValues(folder);
                    if (requestModel != null)
                    {
                        requestModel.token = sharePointToken.access_token;
                        await createRequests.POSTAsync(requestModel);
                        return request.CreateResponse(HttpStatusCode.OK, "Successfully Created Folder " + folder.FolderName);
                    }
                    else
                        throw new Exception();
                }
                else
                    throw new Exception();
            }
            catch(Exception)
            {
                return new HttpResponseMessage(HttpStatusCode.InternalServerError);
            }
        }


        /// <summary>
        /// Description : POST Request to Delete a Particular Folder in SharePoint API
        ///               with the required folder name along with token and digest value that it
        ///               retrieves by internally calling the Get Form Digest Method.
        /// </summary>
        /// <returns>Whether the Folder is deleted successfully or not.</returns>
        [System.Web.Http.Route("api/DeleteFolder")]
        [HttpPost]
        public async Task<object> DeleteFolder(FolderModel folder)
        {
            var configuration = new HttpConfiguration();
            var request = new HttpRequestMessage();
            request.SetConfiguration(configuration);
            try
            {
                SharePointResponse sharePointToken = await GetFormDigest();
                if (sharePointToken != null)
                {
                    RequestModel requestModel = createRequests.DeleteSharePointFolderValues(folder);
                    if (requestModel != null)
                    {
                        requestModel.token = sharePointToken.access_token;
                        requestModel.formDigestValue = sharePointToken.formDigestValue;
                        await createRequests.POSTAsync(requestModel);
                        return request.CreateResponse(HttpStatusCode.OK, "Successfully Deleted Folder " + folder.FolderName);
                    }
                    else
                        throw new Exception();
                }
                else
                    throw new Exception();
            }
            catch (Exception)
            {
                return new HttpResponseMessage(HttpStatusCode.InternalServerError);
            }
        }

        /// <summary>
        /// Description : POST Request to Check whether a Particular Folder in SharePoint API
        ///               with the given folder name is empty or not along with token and digest value that it
        ///               retrieves by internally calling the Get Form Digest Method.
        /// </summary>
        /// <returns>Whether the Folder is Empty or not.</returns>
        [System.Web.Http.Route("api/EmptyFolder")]
        [HttpGet]
        public async Task<object> EmptyFolder(FolderModel folder)
        {
            var configuration = new HttpConfiguration();
            var request = new HttpRequestMessage();
            request.SetConfiguration(configuration);
            try
            {
                folder = (FolderModel)await GetFolderContents(folder);
                if (folder.fileCount < 0 && folder.folderCount < 0)
                    return request.CreateResponse(HttpStatusCode.OK, "Folder " + folder.FolderName + " is Empty");
                else
                    return request.CreateResponse(HttpStatusCode.OK, "Folder " + folder.FolderName + " is not Empty");
            }
            catch (Exception)
            {
                return new HttpResponseMessage(HttpStatusCode.InternalServerError);
            }
        }

        /// <summary>        
        /// Description : POST Request to Retrieve the Contents of a Particular Folder in SharePoint API
        ///               with the given folder name. It also sends the request along with token that it
        ///               retrieves by internally calling the Get Token Method.
        /// </summary>
        /// <returns>The Folder Contents</returns>
        [System.Web.Http.Route("api/foldercontents")]
        [HttpGet]
        public async Task<object> GetFolderContents(FolderModel folder)
        {
            var configuration = new HttpConfiguration();
            var request = new HttpRequestMessage();
            request.SetConfiguration(configuration);
            try
            {
                if (FolderExits(folder) != null)       //Checks if the Folder exists before retrieving contents
                {
                    SharePointResponse sharePointToken = await GetSharepointToken();
                    if (sharePointToken != null)
                    {
                        RequestModel requestModel = createRequests.GetFolderContentSharePointValues(folder);
                        if (requestModel != null)
                        {
                            folder.fileCount = folder.folderCount = 0;
                            folder.files = new List<string>();
                            folder.SubFolders = new List<string>();
                            requestModel.token = sharePointToken.access_token;
                            string responses = (string)await createRequests.GETAsync(requestModel);
                            string[] response = responses.Split('@');
                            string[] fileResult = response[0].Split('\"');
                            string[] folderResult = response[1].Split('\"');
                            for (var i = 0; i < fileResult.Length; i++)
                            {
                                if (fileResult[i].Contains("metadata"))
                                    folder.fileCount++;
                                if (fileResult[i].Contains("ServerRelativeUrl"))
                                {
                                    i += 2;
                                    string[] temp = fileResult[i].Split('/');
                                    folder.files.Add(temp[temp.Length - 1]);
                                }
                            }
                            for (var i = 0; i < folderResult.Length; i++)
                            {
                                if (folderResult[i].Contains("metadata"))
                                    folder.folderCount++;
                                if (folderResult[i].Contains("ServerRelativeUrl"))
                                {
                                    i += 2;
                                    string[] temp = folderResult[i].Split('/');
                                    folder.SubFolders.Add(temp[temp.Length - 1]);
                                }
                            }

                            return folder;
                        }
                        else
                            throw new Exception();
                    }
                    else
                        throw new Exception();
                }
                else
                    throw new Exception();
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Description : GET Request to Check if a Particular Folder in SharePoint Shared Documents 
        ///               with the given folder name exists or not. It also sends the request along with token that it
        ///               retrieves by internally calling the Get Token Method.
        /// </summary>
        /// <returns>The Folder Contents</returns>
        [System.Web.Http.Route("api/folderexists")]
        [HttpGet]
        public async Task<object> FolderExits(FolderModel folder)
        {
            var configuration = new HttpConfiguration();
            var request = new HttpRequestMessage();
            request.SetConfiguration(configuration);
            try
            {
                SharePointResponse sharePointToken = await GetSharepointToken();
                if (sharePointToken != null)
                {
                    RequestModel requestModel = createRequests.FolderExistsSharePointValues(folder);
                    if (requestModel != null)
                    {
                        requestModel.token = sharePointToken.access_token;
                        string responses = (string)await createRequests.GETAsync(requestModel);
                        string[] response = responses.Split('\"');
                        for (var i = 0; i < response.Length; i++)
                        {
                            if (response[i].Contains("ListItemAllFields"))
                                return null;
                        }
                        return responses;
                    }
                    else
                        throw new Exception();
                }
                else
                    throw new Exception();
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Description : POST Request to Upload a Particular File to a particular Folder in SharePoint API
        ///               that is sent along with token and digest value that it retrieves by internally calling the Get Form Digest Method.
        /// </summary>
        /// <returns>Whether the File is Uploaded Successfully or not</returns>
        [System.Web.Http.Route("api/UploadFile")]
        [HttpPost]
        public async Task<object> UploadFile()
        {
            var httpRequest = HttpContext.Current.Request;
            HttpPostedFile FileUpload = httpRequest.Files[0];      //Retrieves the File that is sent with the HTTP Request
            string folderName = httpRequest.Form["Folder_Name"];   // This is used to retrieve from the Request the Destination Folder Name
            
            int FileLen = FileUpload.ContentLength;   //Length of the File sent to be Uploaded
            byte[] input = new byte[FileLen];         //Initiliaze a Byte Array

            // Initialize the stream.
            Stream MyStream = FileUpload.InputStream;

            // Read the file into the byte array.
            MyStream.Read(input, 0, FileLen);

            string path = httpRequest.Form["Path"];

            var configuration = new HttpConfiguration();
            var request = new HttpRequestMessage();
            request.SetConfiguration(configuration);
            try
            {
                SharePointResponse sharePointToken = await GetFormDigest();
                if (sharePointToken != null)
                {
                    FolderModel folder = new FolderModel()
                    {
                        FolderName = folderName,
                        fileName = httpRequest.Files[0].FileName,
                        path = path ?? null
                    };
                    RequestModel requestModel = createRequests.UploadFileSharePointValues(folder);
                    if (requestModel != null)
                    {
                        requestModel.token = sharePointToken.access_token;
                        requestModel.formDigestValue = sharePointToken.formDigestValue;
                        requestModel.httpPostedFile = new ByteArrayContent(input);
                        await createRequests.POSTAsync(requestModel);
                        return request.CreateResponse(HttpStatusCode.OK, "Successfully Uploaded File " + folder.fileName);
                    }
                    else
                        throw new Exception();
                }
                else
                    throw new Exception();
            }
            catch (Exception)
            {
                return new HttpResponseMessage(HttpStatusCode.InternalServerError);
            }
        }


        /// <summary>        
        /// Description : POST Request to Delete a Particular File in SharePoint API
        ///               with the required folder name along with token and digest value that it
        ///               retrieves by internally calling the Get Form Digest Method.
        /// </summary>
        /// <param name="folder">Needs 2 parameters namely 1) The Folder Name and 2) File Name </param>
        /// <returns>Whether the File is deleted successfully or not.</returns>
        [System.Web.Http.Route("api/DeleteFile")]
        [HttpPost]
        public async Task<object> DeleteFile(FolderModel folder)
        {
            var configuration = new HttpConfiguration();
            var request = new HttpRequestMessage();
            request.SetConfiguration(configuration);
            try
            {
                SharePointResponse sharePointToken = await GetFormDigest();
                if (sharePointToken != null)
                {
                    RequestModel requestModel = createRequests.DeleteFileSharePointValues(folder);
                    if (requestModel != null)
                    {
                        requestModel.token = sharePointToken.access_token;
                        requestModel.formDigestValue = sharePointToken.formDigestValue;
                        await createRequests.POSTAsync(requestModel);
                        return request.CreateResponse(HttpStatusCode.OK, "Successfully Deleted File " + folder.fileName);
                    }
                    else
                        throw new Exception();
                }
                else
                    throw new Exception();
            }
            catch (Exception)
            {
                return new HttpResponseMessage(HttpStatusCode.InternalServerError);
            }
        }

        /// <summary>
        /// Description : Gets the Form Digest Value using a Valid Token supplied along 
        ///               with the Request which is retrieved by calling the GetToken Method
        ///               internally
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
                if (sharePointToken != null)
                {
                    RequestModel requestModel = createRequests.CreateSharePointFormDigestRequestValues();
                    if (requestModel != null)
                    {

                        requestModel.token = sharePointToken.access_token;
                        string formDigestResult = (string)await createRequests.POSTAsync(requestModel);
                        string[] formDigest = formDigestResult.Split('\"');
                        for (var i = 0; i < formDigest.Length; i++)
                        {
                            if (formDigest[i].Equals("FormDigestTimeoutSeconds"))
                            {
                                sharePointToken.formDigestTimeout = int.Parse(formDigest[i + 1].Replace(":", "").Replace(",", ""));
                            }
                            if (formDigest[i].Equals("FormDigestValue"))
                            {
                                sharePointToken.formDigestValue = formDigest[i + 2];
                                break;
                            }
                        }
                        return sharePointToken;
                    }
                    else
                        throw new Exception();
                }
                else
                    throw new Exception();
            }
            catch (Exception)
            {
                return null;
            }
        }

    }
}
