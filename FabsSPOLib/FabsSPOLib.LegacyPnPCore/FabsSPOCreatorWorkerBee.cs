using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using System.Net;
using System.Security;
using System.Threading;
using Microsoft.Online.SharePoint.TenantAdministration;


namespace FabsSPOLib.LegacyPnPCore
{
    public class FabsSPOCreatorWorkerBee
    {
        public static void CreateSubWeb(string scurl, string subwebname, string title, string desc)
        {
            string UserName = "fabian@****.onmicrosoft.com"; //Supply your User Name OR this could be done using App Only in Azure AD
            string Password = "*********"; //Supply your own or this would go away if you used Azure AD by registering an App usign Azure AD

            using (ClientContext ctx = new ClientContext(scurl))
            {
                SecureString passWord = new SecureString();
                foreach (char c in Password.ToCharArray()) passWord.AppendChar(c);
                ctx.Credentials = new SharePointOnlineCredentials(UserName, passWord);

                WebCreationInformation wci = new WebCreationInformation();
                wci.Url = subwebname; // This url is relative to the url provided in the context

                wci.Title = title;
                wci.Description = desc;
                wci.UseSamePermissionsAsParentSite = true;
                wci.WebTemplate = "STS#0";
                wci.Language = 1033;

                Web w = ctx.Site.RootWeb.Webs.Add(wci);
                ctx.ExecuteQuery();
            }
        }

        public static void CreateSiteCollection(string scname)
        {
            string TenantURL = "https://fwilliams-admin.sharepoint.com";
            string Title = "Fabian Create Site Collection Demo Alpha";
            string Url = "https://fwilliams.sharepoint.com/sites/" + scname;
            string UserName = "fabian@****.onmicrosoft.com";
            string Password = "***********";

            //Open the Tenant Administration Context with the Tenant Admin Url
            using (ClientContext tenantContext = new ClientContext(TenantURL))
            {
                //Authenticate with a Tenant Administrator
                SecureString passWord = new SecureString();
                foreach (char c in Password.ToCharArray()) passWord.AppendChar(c);
                tenantContext.Credentials = new SharePointOnlineCredentials(UserName, passWord);

                var tenant = new Tenant(tenantContext);

                //Properties of the New SiteCollection
                var siteCreationProperties = new SiteCreationProperties();

                //New SiteCollection Url
                siteCreationProperties.Url = Url;

                //Title of the Root Site
                siteCreationProperties.Title = Title;

                //Email of Owner
                siteCreationProperties.Owner = UserName;

                //Template of the Root Site. Using Team Site for now.
                siteCreationProperties.Template = "COMMUNITY#0";

                //Storage Limit in MB
                siteCreationProperties.StorageMaximumLevel = 100;

                //UserCode Resource Points Allowed
                siteCreationProperties.UserCodeMaximumLevel = 200;

                //Create the SiteCollection
                SpoOperation spo = tenant.CreateSite(siteCreationProperties);

                tenantContext.Load(tenant);

                //We will need the IsComplete property to check if the provisioning of the Site Collection is complete.
                tenantContext.Load(spo, i => i.IsComplete);

                tenantContext.ExecuteQuery();

                //Check if provisioning of the SiteCollection is complete.
                while (!spo.IsComplete)
                {
                    //Wait for 30 seconds and then try again
                    System.Threading.Thread.Sleep(30000);
                    spo.RefreshLoad();
                    tenantContext.ExecuteQuery();
                }

                System.Console.WriteLine("SiteCollection Created.");

            }


        }
    }
}
