using Spectre.Console;

using System.Management;
using System.Text;

namespace SCCMInfo.Enrichers
{
    internal sealed class SmsAdminEnricher : IInstanceEnricher
    {
        public string ClassName => "SMS_Admin";

        public void Enrich(
            ManagementBaseObject targetInstance,
            Table table,
            StringBuilder logMessage,
            ManagementScope scope)
        {
            table.AddRow("Enrichment", "No enrichment implemented yet for SMS_Admin");
            logMessage.AppendLine("SMS_Admin enrichment stub executed");
        }
    }
}
