﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SharePoint_Apps.Models
{
    public class RequestModel
    {
        /// <summary>
        /// Description : This model is used to store the URL that will be used for each POST/GET
        ///               Request, along with body and values wherever needed. Will also be used to
        ///               store the token incase it needs to be sent along with the POST/GET Request
        /// </summary>
        public string URL { get; set; }
        public Dictionary<string,string> Values { get; set; }
        public string body { get; set; }
        public string token { get; set; }
        public string formDigestValue { get; set; }
        public int type { get; set; }
    }
}