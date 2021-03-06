﻿using System;
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
        ///              Library. Also used to Store the File names and filecount, along with folder names
        ///              and folder count when used to retrieve contents of a folder.
        /// </summary>
        public string FolderName { get; set; }
        public List<string> SubFolders { get; set; }
        public string fileName { get; set; }
        public List<string> files { get; set; }      //Used to Store Filenames when uploading Mutiple Files. Also Used to Store the File names under a particular Folder.
        public int fileCount { get; set; }
        public int folderCount { get; set; }
        public int type { get; set; }     //Whether its a Client or Vendor or User
        public string path { get; set; }   //Used to Store Path of the File to be Uploaded. Only used during Upload Operation.
    }
}