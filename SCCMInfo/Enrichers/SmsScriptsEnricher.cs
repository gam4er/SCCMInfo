using Spectre.Console;

using System.Management;
using System.Text;

namespace SCCMInfo.Enrichers
{
    internal sealed class SmsScriptsEnricher : IInstanceEnricher
    {
        public string ClassName => "SMS_Scripts";

        public void Enrich(
            ManagementBaseObject targetInstance,
            Table table,
            StringBuilder logMessage,
            ManagementScope scope)
        {
            table.AddRow("Enrichment", "No enrichment implemented yet for SMS_Scripts");
            logMessage.AppendLine("SMS_Scripts enrichment stub executed");
        }
    }
}
