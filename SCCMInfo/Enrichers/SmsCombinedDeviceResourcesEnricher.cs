using Spectre.Console;

using System.Management;
using System.Text;

namespace SCCMInfo.Enrichers
{
    internal sealed class SmsCombinedDeviceResourcesEnricher : IInstanceEnricher
    {
        public string ClassName => "SMS_CombinedDeviceResources";

        public void Enrich(
            ManagementBaseObject targetInstance,
            Table table,
            StringBuilder logMessage,
            ManagementScope scope)
        {
            table.AddRow("Enrichment", "No enrichment implemented yet for SMS_CombinedDeviceResources");
            logMessage.AppendLine("SMS_CombinedDeviceResources enrichment stub executed");
        }
    }
}
