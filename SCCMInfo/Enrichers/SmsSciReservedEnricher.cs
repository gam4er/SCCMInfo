using Spectre.Console;

using System.Management;
using System.Text;

namespace SCCMInfo.Enrichers
{
    internal sealed class SmsSciReservedEnricher : IInstanceEnricher
    {
        public string ClassName => "SMS_SCI_Reserved";

        public void Enrich(
            ManagementBaseObject targetInstance,
            Table table,
            StringBuilder logMessage,
            ManagementScope scope)
        {
            table.AddRow("Enrichment", "No enrichment implemented yet for SMS_SCI_Reserved");
            logMessage.AppendLine("SMS_SCI_Reserved enrichment stub executed");
        }
    }
}
