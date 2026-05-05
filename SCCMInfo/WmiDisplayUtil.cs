using Spectre.Console;

using Spectre.Console;

using System;
using System.Collections;
using System.Globalization;
using System.Management;
using System.Text;

namespace SCCMInfo
{
    internal static class WmiDisplayUtil
    {
        private const int MaxSerializationDepth = 6;

        public static Table CreateEventTable(string className)
        {
            var table = new Table();
            table.Title("[bold yellow]WMI event[/]");
            table.AddColumn("Property");
            table.AddColumn("Value");

            AddPlainTextRow(table, "ClassName", className ?? string.Empty);
            return table;
        }

        public static void AddPlainTextRow(Table table, string propertyName, string propertyValue)
        {
            if (table == null)
            {
                return;
            }

            table.AddRow(
                new Text(propertyName ?? string.Empty),
                new Text(propertyValue ?? string.Empty));
        }

        public static string GetClassName(ManagementBaseObject managementObject)
        {
            if (managementObject == null)
            {
                return "<null>";
            }

            try
            {
                if (managementObject.ClassPath != null && !string.IsNullOrWhiteSpace(managementObject.ClassPath.ClassName))
                {
                    return managementObject.ClassPath.ClassName;
                }
            }
            catch
            {
            }

            try
            {
                return managementObject["__CLASS"]?.ToString() ?? "<unknown>";
            }
            catch
            {
                return "<unknown>";
            }
        }

        public static string GetPropertyType(PropertyData property)
        {
            if (property == null)
            {
                return "<unknown>";
            }

            try
            {
                string propertyType = property.Type.ToString();

                object propertyValue = null;
                try
                {
                    propertyValue = property.Value;
                }
                catch
                {
                }

                if (propertyValue is Array && propertyType.IndexOf("[]", StringComparison.Ordinal) < 0)
                {
                    return propertyType + "[]";
                }

                return propertyType;
            }
            catch
            {
                try
                {
                    return property.Value?.GetType().FullName ?? "<null>";
                }
                catch
                {
                    return "<unknown>";
                }
            }
        }

        public static string FormatValue(object value)
        {
            var builder = new StringBuilder();
            AppendValue(builder, value, 0);
            return builder.ToString();
        }

        public static void AppendStructuredProperty(
            StringBuilder builder,
            string propertyName,
            string propertyType,
            string propertyValue,
            int indentLevel = 0)
        {
            if (builder == null)
            {
                return;
            }

            string indent = new string(' ', Math.Max(0, indentLevel) * 2);

            builder.AppendLine($"{indent}property name: {propertyName}");
            builder.AppendLine($"{indent}property type: {propertyType}");
            builder.AppendLine($"{indent}property value:");
            AppendIndentedBlock(builder, propertyValue, indentLevel + 1);
        }

        public static void AppendIndentedBlock(StringBuilder builder, string text, int indentLevel)
        {
            if (builder == null)
            {
                return;
            }

            string indent = new string(' ', Math.Max(0, indentLevel) * 2);
            string normalized = text ?? string.Empty;
            string[] lines = normalized
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .Split('\n');

            foreach (string line in lines)
            {
                builder.Append(indent);
                builder.AppendLine(line);
            }
        }

        private static void AppendValue(StringBuilder builder, object value, int depth)
        {
            if (depth >= MaxSerializationDepth)
            {
                builder.Append("<max-depth-reached>");
                return;
            }

            if (value == null)
            {
                builder.Append("<null>");
                return;
            }

            switch (value)
            {
                case string stringValue:
                    builder.Append(stringValue);
                    return;

                case char charValue:
                    builder.Append(charValue);
                    return;

                case bool boolValue:
                    builder.Append(boolValue ? "true" : "false");
                    return;

                case DateTime dateTimeValue:
                    builder.Append(dateTimeValue.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture));
                    return;

                case ManagementBaseObject managementObject:
                    AppendManagementObject(builder, managementObject, depth + 1);
                    return;

                case IEnumerable enumerable when !(value is string):
                    AppendEnumerable(builder, enumerable, depth + 1);
                    return;

                case IFormattable formattable:
                    builder.Append(formattable.ToString(null, CultureInfo.InvariantCulture));
                    return;

                default:
                    builder.Append(value.ToString());
                    return;
            }
        }

        private static void AppendManagementObject(StringBuilder builder, ManagementBaseObject managementObject, int depth)
        {
            builder.Append('{');

            bool hasValues = false;

            foreach (PropertyData property in managementObject.Properties)
            {
                if (hasValues)
                {
                    builder.AppendLine(",");
                }
                else
                {
                    builder.AppendLine();
                }

                AppendIndent(builder, depth);
                builder.Append(property.Name);
                builder.Append(" = ");

                try
                {
                    AppendValue(builder, property.Value, depth);
                }
                catch (Exception ex)
                {
                    builder.Append("<value-read-failed: ");
                    builder.Append(ex.Message);
                    builder.Append('>');
                }

                hasValues = true;
            }

            if (hasValues)
            {
                builder.AppendLine();
                AppendIndent(builder, depth - 1);
            }

            builder.Append('}');
        }

        private static void AppendEnumerable(StringBuilder builder, IEnumerable enumerable, int depth)
        {
            builder.Append('[');

            bool hasValues = false;

            foreach (object item in enumerable)
            {
                if (hasValues)
                {
                    builder.AppendLine(",");
                }
                else
                {
                    builder.AppendLine();
                }

                AppendIndent(builder, depth);
                AppendValue(builder, item, depth);
                hasValues = true;
            }

            if (hasValues)
            {
                builder.AppendLine();
                AppendIndent(builder, depth - 1);
            }

            builder.Append(']');
        }

        private static void AppendIndent(StringBuilder builder, int indentLevel)
        {
            builder.Append(' ', Math.Max(0, indentLevel) * 2);
        }
    }
}