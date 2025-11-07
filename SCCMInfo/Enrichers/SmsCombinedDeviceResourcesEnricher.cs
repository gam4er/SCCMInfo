using Spectre.Console;

using System;
using System.Collections.Generic;
using System.Globalization;
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
            var enrichmentLog = new StringBuilder();
            bool headerAdded = false;

            foreach (PropertyData property in targetInstance.Properties)
            {
                string propertyValue = FormatPropertyValue(property.Value);

                if (string.IsNullOrWhiteSpace(propertyValue))
                {
                    continue;
                }

                if (!headerAdded)
                {
                    table.AddRow("[bold yellow]Enrichment[/]", "[bold yellow]SMS_CombinedDeviceResources non-empty properties[/]");
                    headerAdded = true;
                }

                table.AddRow(property.Name, propertyValue);
                enrichmentLog.AppendLine($"{property.Name.PadRight(30)}\t{propertyValue}");
            }

            if (!headerAdded)
            {
                const string noPropertiesMessage = "SMS_CombinedDeviceResources enrichment: no non-empty properties were found.";
                table.AddRow("Enrichment", noPropertiesMessage);
                logMessage.AppendLine(noPropertiesMessage);
                global::SCCMInfo.SCCMInfo.WriteApplicationEvent(noPropertiesMessage, global::SCCMInfo.SCCMInfo.SmsCombinedDeviceResourcesEventId);
                return;
            }

            string enrichmentHeader = "SMS_CombinedDeviceResources enrichment results:";
            logMessage.AppendLine(enrichmentHeader);
            logMessage.Append(enrichmentLog);

            string eventLogMessage = $"{enrichmentHeader}{Environment.NewLine}{enrichmentLog}";
            global::SCCMInfo.SCCMInfo.WriteApplicationEvent(eventLogMessage, global::SCCMInfo.SCCMInfo.SmsCombinedDeviceResourcesEventId);
        }

        private static string FormatPropertyValue(object value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            if (value is string stringValue)
            {
                return stringValue;
            }

            if (value is Array arrayValue)
            {
                List<string> formattedItems = new List<string>();

                foreach (object item in arrayValue)
                {
                    string formattedItem = FormatPropertyValue(item);
                    if (!string.IsNullOrWhiteSpace(formattedItem))
                    {
                        formattedItems.Add(formattedItem);
                    }
                }

                return string.Join(", ", formattedItems);
            }

            if (value is ManagementBaseObject managementObject)
            {
                return managementObject.ToString() ?? string.Empty;
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }
    }
}
