/*
Copyright (c) Microsoft Corporation
All rights reserved.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. You may obtain a copy of the 
License at http://www.apache.org/licenses/LICENSE-2.0 
    
THIS CODE IS PROVIDED *AS IS* BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING 
WITHOUT LIMITATION ANY IMPLIED WARRANTIES OR CONDITIONS OF TITLE, FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABLITY OR NON-INFRINGEMENT. 

See the Apache Version 2.0 License for specific language governing permissions and limitations under the License.
*/

using mammerla.ServerIntegration;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace mammerla.SharePointIntegration
{
    public class SharePointManager
    {
        private String userId = "test3";
        private String accessTokenSecret = null;
        private EntityTokenManager entityTokenManager;
        private HttpContext context;
        private String sharePointUrl;
        private String accessToken;
        private String requestToken;
        private static SharePointManager current;

        public static SharePointManager Current
        {
            get
            {
                if (current == null)
                {
                    current = new SharePointManager();
                    current.Context = HttpContext.Current;

                    current.Load();
                    
                }

                return current;
            }
        }

        public HttpContext Context
        {
            get
            {
                return this.context;
            }
            set
            {
                this.context = value;
            }
        }

        public String UserId
        {
            get { return this.userId; }
            set
            {
                this.userId = value;
            }
        }

        public string RequestToken
        {
            get 
            {
                return this.requestToken;
            }
            set 
            {
                this.requestToken = value;
            }
        }

        public string AccessToken
        {
            get
            {
                return this.accessToken;
            }

            set 
            {
                this.accessToken = value;
            }
        }

        public String SharePointUrl
        {
            get
            {
                return this.sharePointUrl;
            }

            set 
            {
                this.sharePointUrl = value;
            }
        }

        public string AccessTokenSecret
        {
            get
            {
                return this.accessTokenSecret;
            }
            set 
            {
                this.accessTokenSecret = value;
            }
        }

        public String ConsumerKey
        {
            get
            {
                return ConfigurationManager.AppSettings["ClientId"]; ;
            }
        }

        public String ConsumerSecret
        {
            get
            {
                return ConfigurationManager.AppSettings["ClientSecret"];
            }
        }

        public EntityTokenManager TokenManager
        {
            get
            {
                EntityTokenManager tokenManager = (EntityTokenManager)this.context.Application["SPTokenManager"];

                if (tokenManager == null)
                {
                    string consumerKey = this.ConsumerKey;
                    string consumerSecret = this.ConsumerSecret;

                    if (!string.IsNullOrEmpty(consumerKey))
                    {
                        tokenManager = new EntityTokenManager();
                       
                        this.context.Application["SPTokenManager"] = tokenManager;
                    }
                }

                return (EntityTokenManager)tokenManager;
            }
        }


        public void Load()
        {
            this.RetrieveAccessTokenAndSecret();
        }

        private void RetrieveAccessTokenAndSecret()
        {
            String accessToken;

            EntityTokenManager tokenManager = this.TokenManager;

            if (tokenManager.GetAccessTokenAndSecret(TokenStoreType.SharePoint, this.UserId, out accessToken, out this.accessTokenSecret))
            {
                this.AccessToken = accessToken;
            }
        }
    }
}