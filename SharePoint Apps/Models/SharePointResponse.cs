using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SharePoint_Apps.Models
{
    public class SharePointResponse
    {
        /// <summary>
        /// Description : This Model serves to store the Token Information 
        ///               after successfully logging into Sharepoint using Client ID and
        ///               Client Secret. Also used to Store the Form Digest Information
        ///               namely Form Digest Value and Timeout
        /// </summary>
        public string token_type { get; set; }
        public string expires_in { get; set; }
        public string not_before { get; set; }
        public string expires_on { get; set; }
        public string resource { get; set; }
        public string access_token { get; set; }
        public int formDigestTimeout { get; set; }
        public string formDigestValue { get; set; }


    }
}