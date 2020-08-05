using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SharePoint_Apps.Models
{
    public class FolderModel
    {
        public string FolderName { get; set; }
        public List<string> SubFolders { get; set; }
    }
}