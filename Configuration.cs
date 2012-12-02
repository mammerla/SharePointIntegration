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
using System.Web;
using System.Web.Configuration;
using System.Data.EntityClient;
using System.Configuration;

namespace mammerla.SharePointIntegration
{
    public enum SharePointProxyAuthMode
    {
        OAuth,
        Impersonate
    }

    public static class Configuration
    {
        private static readonly string SharePointProxyAuthModeState = WebConfigurationManager.AppSettings.Get("SharePointProxyAuthMode");
        private static readonly string SharePointProxyAuthModeToMy = WebConfigurationManager.AppSettings.Get("SharePointProxyAuthModeToMy");
        private static SharePointProxyAuthMode? authMode;
        private static SharePointProxyAuthMode? authModeToMy;
        private static bool connectToSocial;
        private static bool doSqlLogging;

        public static String ContextBaseUrl
        {
            get
            {
                return UrlUtilities.GetBaseUrlFromFullUrl(ContextUrl);
            }
        }

        public static String ContextUrl
        {
            get
            {
                return WebConfigurationManager.AppSettings.Get("SharePointContextUrl");
            }
        }
        public static bool LogToSql
        {
            get
            {
                String logToSqlValue = WebConfigurationManager.AppSettings.Get("LogToSql");

                if (logToSqlValue == null)
                {
                    return false;
                }

                bool result = false;

                if (!Boolean.TryParse(logToSqlValue, out result))
                {
                    return false;
                }

                return result;
            }
        }

        public static bool ConnectToSocial
        {
            get
            {
                String socialValue = WebConfigurationManager.AppSettings.Get("ConnectToSocial");

                if (socialValue == null)
                {
                    return false;
                }

                bool result = false;

                if (!Boolean.TryParse(socialValue, out result))
                {
                    return false;
                }

                return result;
            }
        }

        public static SharePointProxyAuthMode AuthMode
        {
            get
            {
                if (authMode == null)
                {
                    String authModeStr = WebConfigurationManager.AppSettings.Get("SharePointProxyAuthMode");

                    if (authModeStr == null)
                    {
                        return SharePointProxyAuthMode.OAuth;
                    }

                    authMode = SharePointProxyAuthMode.OAuth;

                    switch (authModeStr.ToLower())
                    {
                        case "oauth":
                            authMode = SharePointProxyAuthMode.OAuth;
                            break;

                        case "impersonate":
                            authMode = SharePointProxyAuthMode.Impersonate;
                            break;
                    }
                }

                return (SharePointProxyAuthMode)authMode;
            }
        }


        public static SharePointProxyAuthMode AuthModeToMy
        {
            get
            {
                if (authModeToMy == null)
                {
                    String authModeStr = WebConfigurationManager.AppSettings.Get("SharePointProxyAuthModeToMy");

                    if (authModeStr == null)
                    {
                        return SharePointProxyAuthMode.OAuth;
                    }

                    authModeToMy = SharePointProxyAuthMode.OAuth;

                    switch (authModeStr.ToLower())
                    {
                        case "oauth":
                            authModeToMy = SharePointProxyAuthMode.OAuth;
                            break;

                        case "impersonate":
                            authModeToMy = SharePointProxyAuthMode.Impersonate;
                            break;
                    }
                }

                return (SharePointProxyAuthMode)authModeToMy;
            }
        }
    }
}