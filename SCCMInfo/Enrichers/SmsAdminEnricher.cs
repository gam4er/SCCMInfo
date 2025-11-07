using Spectre.Console;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
            if (targetInstance == null)
            {
                table.AddRow("Enrichment", "SMS_Admin enrichment: target instance is null.");
                logMessage.AppendLine("SMS_Admin enrichment failed because target instance was null.");
                return;
            }

            Dictionary<string, object?> properties = targetInstance.Properties
                .Cast<PropertyData>()
                .ToDictionary(property => property.Name, property => property.Value, StringComparer.OrdinalIgnoreCase);

            var processedProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var enrichmentLog = new StringBuilder();
            bool headerAdded = false;

            string [] orderedProperties =
            {
                "AccountType",
                "AdminID",
                "AdminSid",
                "Categories",
                "CategoryNames",
                "CollectionNames",
                "CreatedBy",
                "CreatedDate",
                "DisplayName",
                "DistinguishedName",
                "IsCovered",
                "IsDeleted",
                "IsGroup",
                "LastModifiedBy",
                "LastModifiedDate",
                "LogonName",
                "RoleNames",
                "Roles",
                "SKey",
                "SourceSite"
            };

            foreach (string propertyName in orderedProperties)
            {
                if (!properties.TryGetValue(propertyName, out object? rawValue))
                {
                    continue;
                }

                string formattedValue = FormatSimpleValue(rawValue);
                if (string.IsNullOrWhiteSpace(formattedValue))
                {
                    continue;
                }

                string friendlyName = ToFriendlyName(propertyName);
                AddRow(table, enrichmentLog, ref headerAdded, friendlyName, formattedValue);
                processedProperties.Add(propertyName);
            }

            if (properties.TryGetValue("ExtendedData", out object? extendedDataValue))
            {
                if (AddDetailedCollectionRows(table, enrichmentLog, ref headerAdded, "Extended Data", extendedDataValue))
                {
                    processedProperties.Add("ExtendedData");
                }
            }

            if (properties.TryGetValue("Permissions", out object? permissionsValue))
            {
                if (AddDetailedCollectionRows(table, enrichmentLog, ref headerAdded, "Permissions", permissionsValue))
                {
                    processedProperties.Add("Permissions");
                }
            }

            foreach (KeyValuePair<string, object?> property in properties)
            {
                if (processedProperties.Contains(property.Key) || property.Key.StartsWith("__", StringComparison.Ordinal))
                {
                    continue;
                }

                string formattedValue = FormatSimpleValue(property.Value);
                if (string.IsNullOrWhiteSpace(formattedValue))
                {
                    continue;
                }

                string friendlyName = ToFriendlyName(property.Key);
                AddRow(table, enrichmentLog, ref headerAdded, friendlyName, formattedValue);
            }

            if (headerAdded)
            {
                logMessage.AppendLine("SMS_Admin enrichment results:");
                logMessage.Append(enrichmentLog);
            }
            else
            {
                const string noDataMessage = "SMS_Admin enrichment: no non-empty properties were found.";
                table.AddRow("Enrichment", noDataMessage);
                logMessage.AppendLine(noDataMessage);
            }
        }

        private static bool AddDetailedCollectionRows(
            Table table,
            StringBuilder enrichmentLog,
            ref bool headerAdded,
            string basePropertyName,
            object? rawValue)
        {
            if (rawValue == null)
            {
                return false;
            }

            bool added = false;

            if (rawValue is Array arrayValue)
            {
                int index = 1;

                foreach (object? element in arrayValue)
                {
                    string label = $"{basePropertyName} #{index}";
                    string formatted = FormatSimpleValue(element);

                    if (!string.IsNullOrWhiteSpace(formatted))
                    {
                        AddRow(table, enrichmentLog, ref headerAdded, label, formatted);
                        added = true;
                    }

                    index++;
                }

                return added;
            }

            string fallback = FormatSimpleValue(rawValue);

            if (!string.IsNullOrWhiteSpace(fallback))
            {
                AddRow(table, enrichmentLog, ref headerAdded, basePropertyName, fallback);
                added = true;
            }

            return added;
        }

        private static void AddRow(
            Table table,
            StringBuilder enrichmentLog,
            ref bool headerAdded,
            string propertyName,
            string propertyValue)
        {
            if (string.IsNullOrWhiteSpace(propertyValue))
            {
                return;
            }

            if (!headerAdded)
            {
                table.AddRow("[bold yellow]Enrichment[/]", "[bold yellow]SMS_Admin details[/]");
                headerAdded = true;
            }

            table.AddRow(propertyName, propertyValue);
            AppendToLog(enrichmentLog, propertyName, propertyValue);
        }

        private static void AppendToLog(StringBuilder log, string propertyName, string propertyValue)
        {
            if (string.IsNullOrWhiteSpace(propertyValue))
            {
                return;
            }

            string [] lines = propertyValue
                .Split(new [] { Environment.NewLine }, StringSplitOptions.None)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => line.Trim())
                .ToArray();

            if (lines.Length == 0)
            {
                return;
            }

            log.AppendLine($"{propertyName}: {lines[0]}");

            if (lines.Length == 1)
            {
                return;
            }

            string indent = new string(' ', propertyName.Length + 2);

            for (int index = 1; index < lines.Length; index++)
            {
                log.AppendLine($"{indent}{lines[index]}");
            }
        }

        private static string FormatSimpleValue(object? value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            switch (value)
            {
                case string stringValue:
                    return FormatStringValue(stringValue);

                case string [] stringArray:
                {
                    IEnumerable<string> formattedItems = stringArray
                        .Select(FormatStringValue)
                        .Where(item => !string.IsNullOrWhiteSpace(item));

                    return string.Join(Environment.NewLine, formattedItems);
                }

                case bool boolValue:
                    return boolValue ? "Yes" : "No";

                case DateTime dateTimeValue:
                    return dateTimeValue.ToLocalTime().ToString("u", CultureInfo.InvariantCulture);

                case ManagementBaseObject managementObject:
                    return FormatManagementObjectDetailed(managementObject);

                case Array arrayValue:
                    return FormatArray(arrayValue);

                default:
                    return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            }
        }

        private static string FormatStringValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string trimmed = value.Trim();

            if (LooksLikeDmtfDateTime(trimmed))
            {
                try
                {
                    DateTime convertedDate = ManagementDateTimeConverter.ToDateTime(trimmed);
                    return convertedDate.ToLocalTime().ToString("u", CultureInfo.InvariantCulture);
                }
                catch (FormatException)
                {
                    // Ignore, fall back to the original value.
                }
            }

            string structured = FormatStructuredString(trimmed);
            if (!string.Equals(structured, trimmed, StringComparison.Ordinal))
            {
                return structured;
            }

            return trimmed;
        }

        private static bool LooksLikeDmtfDateTime(string value)
        {
            if (value.Length != 25)
            {
                return false;
            }

            return char.IsDigit(value[0]) && value[14] == '.';
        }

        private static string FormatArray(Array arrayValue)
        {
            List<string> formattedItems = new List<string>();

            foreach (object? item in arrayValue)
            {
                string formatted = FormatSimpleValue(item);
                if (!string.IsNullOrWhiteSpace(formatted))
                {
                    formattedItems.Add(formatted);
                }
            }

            if (formattedItems.Count == 0)
            {
                return string.Empty;
            }

            if (formattedItems.Count == 1)
            {
                return formattedItems[0];
            }

            bool containsNewLine = formattedItems.Any(item => item.IndexOf(Environment.NewLine, StringComparison.Ordinal) >= 0);
            string separator = containsNewLine ? Environment.NewLine + Environment.NewLine : Environment.NewLine;

            return string.Join(separator, formattedItems);
        }

        private static string FormatManagementObjectDetailed(ManagementBaseObject managementObject)
        {
            var builder = new StringBuilder();

            foreach (PropertyData property in managementObject.Properties)
            {
                if (property.Name.StartsWith("__", StringComparison.Ordinal))
                {
                    continue;
                }

                string formattedValue = FormatSimpleValue(property.Value);
                if (string.IsNullOrWhiteSpace(formattedValue))
                {
                    continue;
                }

                string friendlyName = ToFriendlyName(property.Name);

                if (builder.Length > 0)
                {
                    builder.AppendLine();
                }

                builder.Append($"{friendlyName}: {formattedValue}");
            }

            return builder.ToString();
        }

        private static string FormatStructuredString(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string trimmed = value.Trim();
            string [] segments = trimmed
                .Split(new [] { ';', ',', '|' }, StringSplitOptions.RemoveEmptyEntries);

            if (!trimmed.Contains('=') || segments.Length == 0)
            {
                return trimmed;
            }

            var builder = new StringBuilder();

            foreach (string segment in segments)
            {
                string cleanedSegment = segment.Trim();
                if (string.IsNullOrEmpty(cleanedSegment))
                {
                    continue;
                }

                int equalsIndex = cleanedSegment.IndexOf('=');
                if (equalsIndex <= 0 || equalsIndex == cleanedSegment.Length - 1)
                {
                    if (builder.Length > 0)
                    {
                        builder.AppendLine();
                    }

                    builder.Append(cleanedSegment);
                    continue;
                }

                string key = cleanedSegment.Substring(0, equalsIndex).Trim();
                string segmentValue = cleanedSegment.Substring(equalsIndex + 1).Trim();

                if (builder.Length > 0)
                {
                    builder.AppendLine();
                }

                builder.Append($"{ToFriendlyName(key)}: {segmentValue}");
            }

            string formatted = builder.ToString();
            return string.IsNullOrWhiteSpace(formatted) ? trimmed : formatted;
        }

        private static string ToFriendlyName(string propertyName)
        {
            if (string.IsNullOrWhiteSpace(propertyName))
            {
                return propertyName;
            }

            var builder = new StringBuilder(propertyName.Length * 2);
            char previousCharacter = '\0';

            foreach (char current in propertyName)
            {
                if (current == '_')
                {
                    if (builder.Length > 0 && builder[builder.Length - 1] != ' ')
                    {
                        builder.Append(' ');
                    }

                    previousCharacter = current;
                    continue;
                }

                bool isUpperCase = char.IsUpper(current);
                bool isPreviousUpperCase = char.IsUpper(previousCharacter);

                if (builder.Length > 0 && isUpperCase && !isPreviousUpperCase && !char.IsWhiteSpace(previousCharacter))
                {
                    builder.Append(' ');
                }

                builder.Append(current);
                previousCharacter = current;
            }

            return builder.ToString();
        }
    }
}
