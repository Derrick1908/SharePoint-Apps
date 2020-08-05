using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SharePoint_Apps.Models
{
    public class RequestModel
    {
        public string URL { get; set; }
        public Dictionary<string,string> Values { get; set; }
        public string body { get; set; }
        public string token { get; set; }
    }
}