#r "D:\home\site\wwwroot\a_tester_CreateSPSubWeb\bin\FabsSPOLib.LegacyPnPCore.dll"

using FabsSPOLib.LegacyPnPCore;
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
using Newtonsoft.Json;

public class SharePointOnlineSubWebCreate
{
    public string PartitionKey { get; set; }
    public string RowKey { get; set; }
    
    public string SiteRequestedId { get; set;}
    public string SiteCollectionUrl {get; set;}
    public string SiteRequestedName {get; set;}
    public string SiteRequestedTitle {get; set;}
    public string SiteRequestedDescription { get; set;}
}

public static async Task<object> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("SharePoint Online Sub Site Request Function App processed a request.");

    string jsonContent = await req.Content.ReadAsStringAsync();
    var dr = JsonConvert.DeserializeObject<SharePointOnlineSubWebCreate>(jsonContent);
    log.Info($"Site Requested: {dr.SiteRequestedId} with Title {dr.SiteRequestedTitle} at slash: {dr.SiteRequestedName}");
    dr.PartitionKey = "SPSVB2018Demo";
    dr.RowKey = dr.SiteRequestedId;

        if (dr.SiteRequestedId == null || dr.SiteRequestedName == null)
    {
        return req.CreateResponse(HttpStatusCode.BadRequest, new
        {
            error = "Invalid or missing informaiton not supplied!!"
        });
    }
        //Run the code to create the Sub Site
        FabsSPOLib.LegacyPnPCore.FabsSPOCreatorWorkerBee.CreateSubWeb(dr.SiteCollectionUrl,dr.SiteRequestedName,dr.SiteRequestedTitle,dr.SiteRequestedDescription);


        return req.CreateResponse(HttpStatusCode.OK, new {
        message = $"Thank you your Site Request has been Taken."
    });
}
