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
using System.Linq;
using System.Net;
using System.Web;
using System.IO;
using System.Web.Configuration;
using System.Web.SessionState;

namespace mammerla.SharePointIntegration
{

    public class SharePointProxyModule : IHttpHandler, IReadOnlySessionState
    {
        public bool IsReusable
        {
            get
            {
                return true;
            }
        }

        public void ProcessRequest(HttpContext context)
        {
            String subPath = context.Request.Url.ToString();
            bool isClientSvcRequest = false;

            int vtiBinIndex = subPath.IndexOf("_vti_bin");

            if (vtiBinIndex >= 0)
            {
                subPath = subPath.Substring(vtiBinIndex);
            }
            else
            {
                int apiIndex = subPath.IndexOf("_api");

                subPath = subPath.Substring(apiIndex);
            }

            String destinationUrl = context.Request.Headers["x-destination-url"];

            if (destinationUrl == null)
            {
                destinationUrl = context.Request.QueryString["MS.SP.url"];

                if (destinationUrl != null)
                {
                    int vtiBinIndexInDestUrl = destinationUrl.IndexOf("_vti_bin");

                    if (vtiBinIndexInDestUrl >= 0)
                    {
                        subPath = destinationUrl.Substring(vtiBinIndexInDestUrl);

                        if (subPath.IndexOf("client.svc", StringComparison.InvariantCultureIgnoreCase) >= 0)
                        {
                            isClientSvcRequest = true;
                        }
                    }

                    if (vtiBinIndexInDestUrl >= 0)
                    {
                        destinationUrl = destinationUrl.Substring(0, vtiBinIndexInDestUrl - 1);
                    }

                }
            }

            if (destinationUrl == null)
            {
                context.Response.AddHeader("x-sharepointproxy-error", "no.x-destination-url.header.set");
                context.Response.StatusCode = 500;
                return;
            }

            int indexOfVtiBinInSubPath = subPath.IndexOf("_vti_bin/client/", StringComparison.InvariantCultureIgnoreCase);

            if (indexOfVtiBinInSubPath >= 0)
            {

                subPath = subPath.Substring(0, indexOfVtiBinInSubPath) + "_vti_bin/client.svc";
            }

            SharePointProxyAuthMode authMode = Configuration.AuthMode;

            if (Configuration.ContextUrl != null)
            {
                if (UrlUtilities.ServerNamesAreEqual(destinationUrl, Configuration.ContextUrl))
                {
                    authMode = Configuration.AuthModeToMy;
                }
            }

            String url = null;

            if (destinationUrl.IndexOf("_api") >= 0)
            {
                url = UrlUtilities.EnsurePathEndsWithSlash(destinationUrl);
            }
            else
            {
                url = UrlUtilities.EnsurePathEndsWithSlash(destinationUrl) + subPath;
            }

            HttpWebRequest hwr = (HttpWebRequest)WebRequest.Create(url);

            if (authMode == SharePointProxyAuthMode.OAuth)
            {
                object accessToken = null;

                try
                {
                    accessToken = SharePointManager.Current.AccessToken;
                }
                catch (NullReferenceException)
                {
                    ;
                }

                if (accessToken == null)
                {
                    context.Response.AddHeader("x-sharepointproxy-error", "no.oauth.session.token.set.for" + destinationUrl + "|" + Configuration.ContextUrl + "|" + subPath + "|" + UrlUtilities.GetCanonicalServerNameFromFullUrl(destinationUrl) + "|" + UrlUtilities.GetCanonicalServerNameFromFullUrl(Configuration.ContextUrl) + "|" + UrlUtilities.ServerNamesAreEqual(destinationUrl, Configuration.ContextUrl) + "|" + WebConfigurationManager.AppSettings.Get("SharePointProxyAuthModeToMy"));
                    context.Response.StatusCode = 500;
                    return;
                }
                context.Response.AddHeader("x-token-added", (String)accessToken);

                hwr.Headers.Add("Authorization", "Bearer " + (String)accessToken);
            }
            else
            {
                hwr.UseDefaultCredentials = true;
            }

            String[] keys = context.Request.Headers.AllKeys;

            hwr.Method = context.Request.HttpMethod;

            Stream clientRequestStream = context.Request.GetBufferlessInputStream();

            if (hwr.Method != "GET")
            {
                Stream sharePointRequestStream = hwr.GetRequestStream();

                int b = clientRequestStream.ReadByte();
                while (b >= 0)
                {
                    sharePointRequestStream.WriteByte((byte)b);

                    b = clientRequestStream.ReadByte();
                }
            }

            foreach (String key in keys)
            {
                String keyLower = key.ToLower();
                if (key == "Accept")
                {
                    hwr.Accept = context.Request.Headers[key];
                }
                else if (keyLower == "x-requestdigest")
                {
                    hwr.Headers[key] = context.Request.Headers[key];
                }
            }

            hwr.ContentType = context.Request.ContentType;

            WebResponse wresp = null;

            try
            {
                wresp = hwr.GetResponse();
            }
            catch (WebException we)
            {
                context.Response.AddHeader("x-sharepointproxy-error", "request.failed.for" + destinationUrl + "|" + subPath + "|" + UrlUtilities.GetCanonicalServerNameFromFullUrl(destinationUrl) + "|" + we.Message + "|" + authMode.ToString());
                context.Response.StatusCode = 500;
                return;
            }

            Stream sharePointResponseStream = wresp.GetResponseStream();
            Stream clientResponseStream = context.Response.OutputStream;

            context.Response.ContentType = "application/json";

            int br = sharePointResponseStream.ReadByte();

            while (br >= 0)
            {
                clientResponseStream.WriteByte((byte)br);

                br = sharePointResponseStream.ReadByte();
            }
        }
    }
}