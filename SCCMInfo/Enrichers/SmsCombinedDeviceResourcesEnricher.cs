using Spectre.Console;

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
                    global::SCCMInfo.WmiDisplayUtil.AddPlainTextRow(table, "Enrichment", "SMS_CombinedDeviceResources non-empty properties");
                    headerAdded = true;
                }

                global::SCCMInfo.WmiDisplayUtil.AddPlainTextRow(table, property.Name, propertyValue);
                enrichmentLog.AppendLine($"{property.Name.PadRight(30)}\t{propertyValue}");
            }

            if (!headerAdded)
            {
                const string noPropertiesMessage = "SMS_CombinedDeviceResources enrichment: no non-empty properties were found.";
                global::SCCMInfo.WmiDisplayUtil.AddPlainTextRow(table, "Enrichment", noPropertiesMessage);
                logMessage.AppendLine(noPropertiesMessage);
                return;
            }

            logMessage.AppendLine("SMS_CombinedDeviceResources enrichment results:");
            logMessage.Append(enrichmentLog);
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
                var formattedItems = new List<string>();

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
