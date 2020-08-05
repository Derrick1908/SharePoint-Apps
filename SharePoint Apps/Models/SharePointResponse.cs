﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SharePoint_Apps.Models
{
    public class SharePointResponse
    {
        public string token_type { get; set; }
        public string expires_in { get; set; }
        public string not_before { get; set; }
        public string expires_on { get; set; }
        public string resource { get; set; }
        public string access_token { get; set; }

    }
}