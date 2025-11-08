using Spectre.Console;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
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
            try
            {
                AppendAvailability(targetInstance, table, logMessage);
                AppendFileType(targetInstance, table, logMessage);
                AppendArrayProperty(targetInstance, table, logMessage, "AccountUsage");
                AppendEmbeddedProperty(table, logMessage, "PropLists", targetInstance["PropLists"]);
                AppendEmbeddedProperty(table, logMessage, "Props", targetInstance["Props"]);

                logMessage.AppendLine("SMS_SCI_Reserved enrichment completed.");
            }
            catch (Exception ex)
            {
                table.AddRow("EnrichmentError", ex.Message);
                logMessage.AppendLine($"SMS_SCI_Reserved enrichment failed: {ex}");
            }
        }

        private static void AppendAvailability(
            ManagementBaseObject targetInstance,
            Table table,
            StringBuilder logMessage)
        {
            if (targetInstance["Availability"] is uint availability)
            {
                if (AvailabilityMap.TryGetValue(availability, out string? description))
                {
                    table.AddRow("AvailabilityDescription", description);
                    logMessage.AppendLine($"AvailabilityDescription: {description}");
                }
                else
                {
                    table.AddRow("AvailabilityDescription", $"Unknown ({availability})");
                    logMessage.AppendLine($"AvailabilityDescription: Unknown ({availability})");
                }
            }
        }

        private static void AppendFileType(
            ManagementBaseObject targetInstance,
            Table table,
            StringBuilder logMessage)
        {
            if (targetInstance["FileType"] is uint fileType)
            {
                if (FileTypeMap.TryGetValue(fileType, out string? description))
                {
                    table.AddRow("FileTypeDescription", description);
                    logMessage.AppendLine($"FileTypeDescription: {description}");
                }
                else
                {
                    table.AddRow("FileTypeDescription", $"Unknown ({fileType})");
                    logMessage.AppendLine($"FileTypeDescription: Unknown ({fileType})");
                }
            }
        }

        private static void AppendArrayProperty(
            ManagementBaseObject targetInstance,
            Table table,
            StringBuilder logMessage,
            string propertyName)
        {
            if (targetInstance[propertyName] is Array array)
            {
                string serialized = SerializeToJsonLike(array);
                table.AddRow($"{propertyName} (detailed)", serialized);
                logMessage.AppendLine($"{propertyName} (detailed):");
                logMessage.AppendLine(serialized);
            }
        }

        private static void AppendEmbeddedProperty(
            Table table,
            StringBuilder logMessage,
            string propertyName,
            object? value)
        {
            string serialized = SerializeToJsonLike(value);
            table.AddRow($"{propertyName} (detailed)", serialized);
            logMessage.AppendLine($"{propertyName} (detailed):");
            logMessage.AppendLine(serialized);
        }

        private static string SerializeToJsonLike(object? value)
        {
            var builder = new StringBuilder();
            SerializeValue(builder, value, 0);
            return builder.ToString();
        }

        private static void SerializeValue(StringBuilder builder, object? value, int indentLevel)
        {
            if (value is null)
            {
                builder.Append("null");
                return;
            }

            switch (value)
            {
                case string s:
                    builder.Append('"').Append(EscapeString(s)).Append('"');
                    return;
                case char c:
                    builder.Append('"').Append(EscapeString(c.ToString())).Append('"');
                    return;
                case bool b:
                    builder.Append(b ? "true" : "false");
                    return;
                case sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal:
                    if (value is IFormattable formattable)
                    {
                        builder.Append(formattable.ToString(null, CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        builder.Append(value.ToString());
                    }

                    return;
                case ManagementBaseObject managementBaseObject:
                    SerializeManagementObject(builder, managementBaseObject, indentLevel);
                    return;
                case IEnumerable enumerable:
                    SerializeEnumerable(builder, enumerable, indentLevel);
                    return;
                default:
                    builder.Append('"').Append(EscapeString(value.ToString() ?? string.Empty)).Append('"');
                    return;
            }
        }

        private static void SerializeManagementObject(
            StringBuilder builder,
            ManagementBaseObject managementBaseObject,
            int indentLevel)
        {
            builder.Append('{');
            bool hasProperties = false;

            foreach (PropertyData property in managementBaseObject.Properties)
            {
                if (hasProperties)
                {
                    builder.AppendLine(",");
                }
                else
                {
                    builder.AppendLine();
                }

                AppendIndent(builder, indentLevel + 1);
                builder.Append('"').Append(property.Name).Append('"');
                builder.Append(": ");
                SerializeValue(builder, property.Value, indentLevel + 1);
                hasProperties = true;
            }

            if (hasProperties)
            {
                builder.AppendLine();
                AppendIndent(builder, indentLevel);
            }

            builder.Append('}');
        }

        private static void SerializeEnumerable(
            StringBuilder builder,
            IEnumerable enumerable,
            int indentLevel)
        {
            builder.Append('[');
            bool hasElements = false;

            foreach (object? item in enumerable)
            {
                if (hasElements)
                {
                    builder.AppendLine(",");
                }
                else
                {
                    builder.AppendLine();
                }

                AppendIndent(builder, indentLevel + 1);
                SerializeValue(builder, item, indentLevel + 1);
                hasElements = true;
            }

            if (hasElements)
            {
                builder.AppendLine();
                AppendIndent(builder, indentLevel);
            }

            builder.Append(']');
        }

        private static void AppendIndent(StringBuilder builder, int indentLevel)
        {
            builder.Append(' ', Math.Max(0, indentLevel) * 2);
        }

        private static string EscapeString(string value)
        {
            var builder = new StringBuilder(value.Length);

            foreach (char c in value)
            {
                switch (c)
                {
                    case '\\':
                        builder.Append(@"\\");
                        break;
                    case '"':
                        builder.Append("\\\"");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\t':
                        builder.Append("\\t");
                        break;
                    default:
                        if (char.IsControl(c))
                        {
                            builder.AppendFormat(CultureInfo.InvariantCulture, "\\u{0:X4}", (int)c);
                        }
                        else
                        {
                            builder.Append(c);
                        }

                        break;
                }
            }

            return builder.ToString();
        }

        private static readonly IReadOnlyDictionary<uint, string> AvailabilityMap =
            new Dictionary<uint, string>
            {
                { 0, "LOCAL" },
                { 1, "GLOBAL" }
            };

        private static readonly IReadOnlyDictionary<uint, string> FileTypeMap =
            new Dictionary<uint, string>
            {
                { 0, "EMPTY" },
                { 1, "ACTUAL" },
                { 2, "PROPOSED" },
                { 4, "TRANSACTIONS" },
                { 6, "LOCAL_TRANSACTIONS" }
            };
    }
}
