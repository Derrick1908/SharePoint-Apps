using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SharePoint_Apps.Models
{
    public class FolderModel
    {
        /// <summary>
        /// Description: This Model is used for supplying Folder Name and Sub Folders 
        ///              within that Folder to be created in SharePoint under Shared Documents
        ///              Library
        /// </summary>
        public string FolderName { get; set; }
        public List<string> SubFolders { get; set; }
        public string fileName { get; set; }
    }
}