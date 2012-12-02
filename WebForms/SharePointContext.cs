/*
Copyright (c) Microsoft Corporation
All rights reserved.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. You may obtain a copy of the 
License at http://www.apache.org/licenses/LICENSE-2.0 
    
THIS CODE IS PROVIDED *AS IS* BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING 
WITHOUT LIMITATION ANY IMPLIED WARRANTIES OR CONDITIONS OF TITLE, FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABLITY OR NON-INFRINGEMENT. 

See the Apache Version 2.0 License for specific language governing permissions and limitations under the License.
*/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using Microsoft.IdentityModel.S2S.Tokens;
using System.Net;
using System.IO;
using System.Web.Configuration;
using mammerla.ServerIntegration;

namespace mammerla.SharePointIntegration.WebForms
{
    [DefaultProperty("Text")]
    [ToolboxData("<{0}:SharePointContext runat=server></{0}:SharePointContext>")]
    public class SharePointContext : WebControl
    {
        private static string unformattedContextUrl = WebConfigurationManager.AppSettings.Get("SharePointContextUrl");

        String userId = null;

        protected override void RenderContents(HtmlTextWriter output)
        {
            TokenHelper.TrustAllCertificates();
            String pageRequestQueryString = this.Page.Request.QueryString["SPHostUrl"];

            SharePointManager.Current.SharePointUrl = pageRequestQueryString;

            string contextTokenString = TokenHelper.GetContextTokenFromRequest(this.Page.Request);

            if (contextTokenString != null)
            {
                SharePointManager.Current.RequestToken = contextTokenString;

                SharePointContextToken contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, this.Page.Request.Url.Authority);

                if (contextToken.CacheKey == null)
                {
                    throw new Exception("No cache key specified.");
                }
                else
                {
                    userId = contextToken.CacheKey;

                    int lastSlash = userId.LastIndexOf("\\");

                    if (lastSlash >= 0)
                    {
                        userId = userId.Substring(lastSlash + 1, userId.Length - (lastSlash + 1));
                    }
                }

                SharePointManager.Current.TokenManager.StoreNewToken(TokenType.RequestToken, TokenStoreType.SharePoint, this.userId, contextTokenString, "");

                Uri sharepointUrl = new Uri(pageRequestQueryString);

                string accessToken = TokenHelper.GetAccessToken(contextToken, sharepointUrl.Authority).AccessToken;

                SharePointManager.Current.AccessToken = accessToken;
                SharePointManager.Current.TokenManager.ExpireRequestTokenAndStoreNewAccessToken(SharePointManager.Current.ConsumerSecret, SharePointManager.Current.RequestToken, TokenStoreType.SharePoint,this.userId, SharePointManager.Current.AccessToken, String.Empty);

                // this.Context.Session["OAT" + UrlUtilities.CanonicalizeUrlForCompare(sharepointUrl.ToString())] = accessToken;
            }

            output.Write(String.Format(@"
<script language='javascript'>
    var userId = ""{0}"";
", userId));

            if (Configuration.ConnectToSocial)
            {
                RequestUserInformation(output);
            }

            output.Write("</script>");
        }

        private void RequestUserInformation(HtmlTextWriter output)
        {
            HttpWebRequest hwr = (HttpWebRequest)WebRequest.Create(UrlUtilities.EnsurePathEndsWithSlash(GetFormattedContextUrl()) + "_api/SP.Microfeed.MicrofeedManager/");
            hwr.Method = "GET";
            hwr.Accept = "application/json";

            if (Configuration.AuthModeToMy == SharePointProxyAuthMode.Impersonate)
            {
                hwr.UseDefaultCredentials = true;
            }

            WebResponse wresp = null;

            try
            {
                wresp = hwr.GetResponse();
            }           
            catch (WebException exception)
            {
                // todo: better error handling.
                throw exception;
            }
            catch (Microsoft.SharePoint.Client.ServerUnauthorizedAccessException exception)
            {
                // todo: better error handling.
                throw exception;
            }

            Stream sharePointResponseStream = wresp.GetResponseStream();

            StreamReader reader = new StreamReader(sharePointResponseStream);

            StringBuilder sb = new StringBuilder();
            String line = reader.ReadLine();

            sb.AppendLine("var userInfo = " + line.Substring(1));

            while (!reader.EndOfStream)
            {
                sb.AppendLine(line);
                line = reader.ReadLine();
            }

            sb.Append("};");

            String outputA = sb.ToString().Replace("\r", "").Replace("\n", "");

            output.Write(outputA);
        }

        private string GetFormattedContextUrl()
        {
            if (Configuration.AuthModeToMy == SharePointProxyAuthMode.Impersonate)
            {
                return String.Format(unformattedContextUrl, "pkmacct");
            }
            else if (!String.IsNullOrEmpty(this.userId))
            {
                return String.Format(unformattedContextUrl, this.userId);
            }
            else
            {
                return null; 
            }
        }
    }
}
