using System;
using System.Text;
using Microsoft.SharePoint.Client;
using PnPAuthenticationManager = OfficeDevPnP.Core.AuthenticationManager;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Online.SharePoint.TenantAdministration;

namespace AzureFunctionSPSiteAudit
{
    public static class SPSiteAudit
    {
        //Timer Trigger runs at 9pm every day
        [FunctionName("SPSiteAudit")]
        public static void Run([TimerTrigger("0 0 23 * * *")]TimerInfo myTimer, TraceWriter log)
        {
            string userName = System.Environment.GetEnvironmentVariable("SPUser", EnvironmentVariableTarget.Process);
            string password = System.Environment.GetEnvironmentVariable("SPPwd", EnvironmentVariableTarget.Process);
            string SPRootURL = System.Environment.GetEnvironmentVariable("SPRootURL", EnvironmentVariableTarget.Process);
            string sendReportTo = System.Environment.GetEnvironmentVariable("SendReportTo", EnvironmentVariableTarget.Process);
            var authenticationManager = new PnPAuthenticationManager();
            var clientContext = authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(SPRootURL, userName, password);
            try
            {
                log.Info($"SPAudit Timer trigger function executed at: {DateTime.Now}");


                var tenant = new Tenant(clientContext);
                var siteProperties = tenant.GetSiteProperties(0, true);
                clientContext.Load(siteProperties);
                clientContext.ExecuteQuery();

                var sbOutput = new StringBuilder();
                sbOutput.AppendLine("Audit Report Results<br/>");
                sbOutput.AppendLine("---------------------<br/>");
                foreach (SiteProperties sp in siteProperties)
                {
                    //since we're iterating through the site collections we need a new context object for each site collection
                    var siteClientContext = authenticationManager.GetSharePointOnlineAuthenticatedContextTenant(sp.Url, userName, password);

                    Web web = siteClientContext.Web;
                    GroupCollection groupCollection = web.SiteGroups;
                    siteClientContext.Load(groupCollection, gc => gc.Include(group => group.AllowMembersEditMembership, Group => Group.Title));
                    siteClientContext.ExecuteQuery();

                    foreach (Group group in groupCollection)
                    {
                        if (group.AllowMembersEditMembership)
                        {
                            sbOutput.AppendLine($"Site Collection = {sp.Url} . Group with Edit Membership found = {group.Title}<br/>");
                        }
                    }
                }
                sbOutput.AppendLine("---------------------<br/>");
                sbOutput.AppendLine($"Audit Report Complete at {DateTime.Now}");
                OfficeDevPnP.Core.Utilities.MailUtility.SendEmail(clientContext, new string[] { sendReportTo }, null, "SharePoint Audit Report Results", sbOutput.ToString());
                log.Info($"SPAudit completed normally at: {DateTime.Now}");
            }
            catch (Exception ex)
            {

                log.Info($"SPAudit Timer Exception at: {DateTime.Now}");
                log.Info($"Message: {ex.Message}");
                log.Info($"Stack Trace: {ex.StackTrace}");
                OfficeDevPnP.Core.Utilities.MailUtility.SendEmail(clientContext, new string[] { sendReportTo }, null, "SharePoint Audit Report Has Errored", ex.Message + ex.StackTrace);
            }
        }
    }
}