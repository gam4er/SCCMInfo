using Spectre.Console;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.ServiceProcess;
using System.Text;

using SCCMInfo.Enrichers;

namespace SCCMInfo
{
    internal class SCCMInfo
    {
        private const string ServiceName = "SCCMInfo";
        private const string ServiceDisplayName = "SCCM Info";

        public static ManagementScope scope = new ManagementScope();

        private const string ApplicationLogName = "Application";
        private const string EventSourceName = "SCCMInfo";
        private const int EventLogEntryId = 2001;
        private const int EventLogMessageMaxLength = 30000;

        private static readonly IInstanceEnricher[] InstanceEnrichers =
        {
            new SmsDeploymentInfoEnricher(),
            new SmsCombinedDeviceResourcesEnricher(),
            new SmsAdminEnricher(),
            new SmsScriptsEnricher(),
            new SmsSciReservedEnricher()
        };

        private static readonly List<ManagementEventWatcher> ActiveWatchers = new List<ManagementEventWatcher>();

        private static bool IsServiceMode { get; set; }

        private static void ProcMon()
        {
            StopMonitoring();
            WriteLog("monitor", "process monitor start requested");

            string[] classesToMonitor =
            {
                "Win32_Process"
            };

            try
            {
                var localScope = new ManagementScope();
                WriteLog("monitor", $"process monitor scope: {localScope.Path?.Path ?? "<null>"}");

                foreach (string className in classesToMonitor)
                {
                    var query = new WqlEventQuery(
                        "__InstanceCreationEvent",
                        new TimeSpan(0, 0, 1),
                        $"TargetInstance ISA '{className}'");

                    var watcher = new ManagementEventWatcher(localScope, query);
                    watcher.EventArrived += new EventArrivedEventHandler(HandleEvent);
                    watcher.Start();
                    ActiveWatchers.Add(watcher);

                    WriteLog("monitor", $"watcher started: class={className}; query={query.QueryString}");
                }
            }
            catch (Exception ex)
            {
                WriteLog("monitor", $"process monitor start failed{Environment.NewLine}{ex}");
            }
        }

        private static void CCMMon()
        {
            StopMonitoring();
            WriteLog("monitor", "ccm monitor start requested");

            string[] classesToMonitor =
            {
                "SMS_DeploymentInfo",
                "SMS_CombinedDeviceResources",
                "SMS_Admin",
                "SMS_Scripts",
                "SMS_SCI_Reserved"
            };

            try
            {
                scope = WmiUtil.NewWmiConnection();
                if (scope == null)
                {
                    WriteLog("monitor", "ccm monitor aborted because WMI connection is null");
                    return;
                }

                WriteLog("monitor", $"ccm monitor scope: {scope.Path?.Path ?? "<null>"}");

                foreach (string className in classesToMonitor)
                {
                    var query = new WqlEventQuery(
                        "__InstanceCreationEvent",
                        new TimeSpan(0, 0, 1),
                        $"TargetInstance ISA '{className}'");

                    var watcher = new ManagementEventWatcher(scope, query);
                    watcher.EventArrived += new EventArrivedEventHandler(HandleEvent);
                    watcher.Start();
                    ActiveWatchers.Add(watcher);

                    WriteLog("monitor", $"watcher started: class={className}; query={query.QueryString}");
                }
            }
            catch (Exception ex)
            {
                WriteLog("monitor", $"ccm monitor start failed{Environment.NewLine}{ex}");
            }
        }

        static void Main(string[] args)
        {
            args = args ?? Array.Empty<string>();

            WriteLog("startup", $"application start; args={string.Join(" ", args)}");

            if (args.Any(a => string.Equals(a, "--install", StringComparison.OrdinalIgnoreCase)))
            {
                InstallService();
                return;
            }

            bool runAsService = args.Any(a => string.Equals(a, "--service", StringComparison.OrdinalIgnoreCase));
            IsServiceMode = runAsService || !Environment.UserInteractive;

            WriteLog("startup", $"service mode={IsServiceMode}");

            if (IsServiceMode)
            {
                ServiceBase.Run(new SCCMInfoServiceHost());
                return;
            }

            Console.CancelKeyPress += (sender, eventArgs) =>
            {
                WriteLog("shutdown", "console cancel requested");
                StopMonitoring();
            };

            CCMMon();
            Console.ReadLine();
            StopMonitoring();
        }

        private static void InstallService()
        {
            if (!IsRunningOnWindows())
            {
                Console.WriteLine("Установка сервиса поддерживается только в Windows.");
                WriteLog("service", "service installation aborted because current platform is not Windows");
                return;
            }

            try
            {
                if (IsServiceInstalled(ServiceName))
                {
                    Console.WriteLine($"Сервис \"{ServiceName}\" уже установлен.");
                    WriteLog("service", $"service installation skipped because service '{ServiceName}' already exists");
                    return;
                }

                string executablePath = Process.GetCurrentProcess().MainModule.FileName;
                string arguments = $"create \"{ServiceName}\" binPath= \"\\\"{executablePath}\\\" --service\" start= auto DisplayName= \"{ServiceDisplayName}\"";

                var startInfo = new ProcessStartInfo("sc.exe", arguments)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (Process process = Process.Start(startInfo))
                {
                    if (process == null)
                    {
                        Console.WriteLine("Не удалось запустить sc.exe для установки сервиса.");
                        WriteLog("service", "service installation failed because sc.exe process was not created");
                        return;
                    }

                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    WriteLog(
                        "service",
                        $"service installation command completed; exitCode={process.ExitCode}{Environment.NewLine}stdout:{Environment.NewLine}{output}{Environment.NewLine}stderr:{Environment.NewLine}{error}");

                    if (process.ExitCode == 0)
                    {
                        Console.WriteLine("Сервис успешно установлен.");
                        if (!string.IsNullOrWhiteSpace(output))
                        {
                            Console.WriteLine(output.Trim());
                        }

                        WriteLog("service", "service installed successfully");
                    }
                    else
                    {
                        Console.WriteLine("Не удалось установить сервис. Подробности ниже:");
                        if (!string.IsNullOrWhiteSpace(error))
                        {
                            Console.WriteLine(error.Trim());
                        }

                        if (!string.IsNullOrWhiteSpace(output))
                        {
                            Console.WriteLine(output.Trim());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка установки сервиса: {ex.Message}");
                WriteLog("service", $"service installation failed{Environment.NewLine}{ex}");
            }
        }

        private static bool IsServiceInstalled(string serviceName)
        {
            try
            {
                bool installed = ServiceController.GetServices()
                    .Any(service => string.Equals(service.ServiceName, serviceName, StringComparison.OrdinalIgnoreCase));

                WriteLog("service", $"service installed check: name={serviceName}; installed={installed}");
                return installed;
            }
            catch (Exception ex)
            {
                WriteLog("service", $"service installed check failed for '{serviceName}'{Environment.NewLine}{ex}");
                return false;
            }
        }

        private static bool IsRunningOnWindows()
        {
            PlatformID platform = Environment.OSVersion.Platform;
            return platform == PlatformID.Win32NT
                || platform == PlatformID.Win32S
                || platform == PlatformID.Win32Windows
                || platform == PlatformID.WinCE;
        }

        private static void StopMonitoring()
        {
            foreach (ManagementEventWatcher watcher in ActiveWatchers.ToList())
            {
                try
                {
                    watcher.EventArrived -= new EventArrivedEventHandler(HandleEvent);
                    watcher.Stop();
                    watcher.Dispose();
                    WriteLog("monitor", "watcher stopped successfully");
                }
                catch (Exception ex)
                {
                    WriteLog("monitor", $"watcher stop failed{Environment.NewLine}{ex}");
                }
                finally
                {
                    ActiveWatchers.Remove(watcher);
                }
            }
        }

        private static void HandleEvent(object sender, EventArrivedEventArgs e)
        {
            var captureLog = new StringBuilder();
            Table table = null;
            ManagementBaseObject targetInstance = null;
            string className = "<unknown>";
            int propertyCount = 0;
            bool captureSucceeded = false;
            bool enrichmentAttempted = false;
            bool enrichmentSucceeded = false;

            captureLog.AppendLine("capture event start");
            captureLog.AppendLine($"capture event timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
            captureLog.AppendLine($"capture event sender type: {sender?.GetType().FullName ?? "<null>"}");

            try
            {
                if (!TryGetTargetInstance(e, captureLog, out targetInstance))
                {
                    captureLog.AppendLine("capture event aborted because target instance is unavailable");
                    return;
                }

                className = WmiDisplayUtil.GetClassName(targetInstance);
                table = WmiDisplayUtil.CreateEventTable(className);

                captureLog.AppendLine($"capture event class: {className}");
                captureLog.AppendLine("capture event target instance properties:");

                foreach (PropertyData property in targetInstance.Properties)
                {
                    propertyCount++;

                    try
                    {
                        string propertyName = property?.Name ?? "<unknown>";
                        string propertyType = WmiDisplayUtil.GetPropertyType(property);

                        object rawValue = null;
                        try
                        {
                            rawValue = property?.Value;
                        }
                        catch (Exception valueEx)
                        {
                            rawValue = $"<value-read-failed: {valueEx.Message}>";
                        }

                        string propertyValue = WmiDisplayUtil.FormatValue(rawValue);

                        WmiDisplayUtil.AppendStructuredProperty(captureLog, propertyName, propertyType, propertyValue, 1);
                        WmiDisplayUtil.AddPlainTextRow(table, propertyName, propertyValue);
                    }
                    catch (Exception propertyEx)
                    {
                        string failedPropertyName = property?.Name ?? "<unknown>";
                        captureLog.AppendLine($"  capture event property failed: {failedPropertyName}");
                        WmiDisplayUtil.AppendIndentedBlock(captureLog, propertyEx.ToString(), 2);
                        WmiDisplayUtil.AddPlainTextRow(table, failedPropertyName, $"Property capture failed: {propertyEx.Message}");
                    }
                }

                captureSucceeded = true;

                foreach (IInstanceEnricher enricher in InstanceEnrichers)
                {
                    if (!string.Equals(enricher.ClassName, className, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    enrichmentAttempted = true;
                    captureLog.AppendLine($"enrich start: {enricher.GetType().FullName}");

                    try
                    {
                        enricher.Enrich(targetInstance, table, captureLog, scope);
                        enrichmentSucceeded = true;
                        captureLog.AppendLine($"enrich completed: {enricher.GetType().FullName}");
                    }
                    catch (Exception enrichEx)
                    {
                        captureLog.AppendLine($"enrich failed: {enricher.GetType().FullName}");
                        WmiDisplayUtil.AppendIndentedBlock(captureLog, enrichEx.ToString(), 1);
                        WmiDisplayUtil.AddPlainTextRow(table, "EnrichmentError", enrichEx.ToString());
                    }

                    break;
                }

                if (!enrichmentAttempted)
                {
                    captureLog.AppendLine($"enrich skipped: no enricher registered for class {className}");
                }
            }
            catch (Exception ex)
            {
                captureLog.AppendLine("capture event fatal error:");
                WmiDisplayUtil.AppendIndentedBlock(captureLog, ex.ToString(), 1);

                if (table != null)
                {
                    WmiDisplayUtil.AddPlainTextRow(table, "CaptureError", ex.ToString());
                }
            }
            finally
            {
                captureLog.AppendLine($"capture event property count: {propertyCount}");

                WriteLog("capture event", captureLog.ToString());
                WriteEventCaptureSummary(className, propertyCount, captureSucceeded, enrichmentAttempted, enrichmentSucceeded);

                if (!IsServiceMode && table != null)
                {
                    TryRenderTable(table, className);
                }
            }
        }

        private static bool TryGetTargetInstance(
            EventArrivedEventArgs eventArgs,
            StringBuilder captureLog,
            out ManagementBaseObject targetInstance)
        {
            targetInstance = null;

            if (eventArgs == null)
            {
                captureLog.AppendLine("capture event args are null");
                return false;
            }

            ManagementBaseObject newEvent = null;

            try
            {
                newEvent = eventArgs.NewEvent;
            }
            catch (Exception ex)
            {
                captureLog.AppendLine("capture event failed while reading NewEvent");
                WmiDisplayUtil.AppendIndentedBlock(captureLog, ex.ToString(), 1);
                return false;
            }

            if (newEvent == null)
            {
                captureLog.AppendLine("capture event NewEvent is null");
                return false;
            }

            captureLog.AppendLine("capture event envelope:");
            WmiDisplayUtil.AppendIndentedBlock(captureLog, WmiDisplayUtil.FormatValue(newEvent), 1);

            object rawTargetInstance = null;

            try
            {
                rawTargetInstance = newEvent["TargetInstance"];
            }
            catch (Exception ex)
            {
                captureLog.AppendLine("capture event failed while reading TargetInstance");
                WmiDisplayUtil.AppendIndentedBlock(captureLog, ex.ToString(), 1);
                return false;
            }

            if (rawTargetInstance == null)
            {
                captureLog.AppendLine("capture event TargetInstance is null");
                return false;
            }

            targetInstance = rawTargetInstance as ManagementBaseObject;
            if (targetInstance == null)
            {
                captureLog.AppendLine($"capture event TargetInstance has unexpected type: {rawTargetInstance.GetType().FullName}");
                WmiDisplayUtil.AppendIndentedBlock(captureLog, WmiDisplayUtil.FormatValue(rawTargetInstance), 1);
                return false;
            }

            return true;
        }

        private static void TryRenderTable(Table table, string className)
        {
            try
            {
                AnsiConsole.Write(table);
            }
            catch (Exception ex)
            {
                WriteLog("render table", $"table render failed for class {className}{Environment.NewLine}{ex}");
                Console.WriteLine($"Не удалось отрисовать таблицу для {className}. Полные данные записаны в ProcessInfoLog.txt.");
            }
        }

        private static void WriteEventCaptureSummary(
            string className,
            int propertyCount,
            bool captureSucceeded,
            bool enrichmentAttempted,
            bool enrichmentSucceeded)
        {
            string enrichState;

            if (!enrichmentAttempted)
            {
                enrichState = "skipped";
            }
            else
            {
                enrichState = enrichmentSucceeded ? "success" : "failed";
            }

            string summary =
                $"capture event summary: class={className}; properties={propertyCount}; capture={(captureSucceeded ? "success" : "failed")}; enrich={enrichState}; serviceMode={IsServiceMode}";

            WriteLog("write event", $"application event summary prepared: {summary}");
            WriteApplicationEvent(summary);
        }

        public static void WriteLog(string message)
        {
            WriteLog("general", message);
        }

        public static void WriteLog(string keyword, string message)
        {
            string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ProcessInfoLog.txt");

            try
            {
                string logDirectory = Path.GetDirectoryName(logFilePath);

                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                string structuredEntry = FormatStructuredLogEntry(keyword, message);

                using (var writer = new StreamWriter(logFilePath, true, Encoding.UTF8))
                {
                    writer.Write(structuredEntry);
                }
            }
            catch (Exception ex)
            {
                if (!IsServiceMode)
                {
                    Console.WriteLine($"Log write error: {ex.Message}");
                    Console.WriteLine(ex.StackTrace);
                }
            }
        }

        private static string FormatStructuredLogEntry(string keyword, string message)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string normalizedKeyword = string.IsNullOrWhiteSpace(keyword) ? "general" : keyword.Trim();
            string normalizedMessage = message ?? string.Empty;

            string[] lines = normalizedMessage
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .Split('\n');

            var builder = new StringBuilder();

            foreach (string line in lines)
            {
                builder.Append(timestamp);
                builder.Append(" | ");
                builder.Append(normalizedKeyword);
                builder.Append(" | ");
                builder.AppendLine(line);
            }

            return builder.ToString();
        }

        internal static void WriteApplicationEvent(string message, int eventId = EventLogEntryId)
        {
            if (!IsApplicationLogWritable(out string failureReason))
            {
                if (!string.IsNullOrWhiteSpace(failureReason))
                {
                    WriteLog("write event", $"application event log is not writable: {failureReason}");
                }

                return;
            }

            try
            {
                string normalizedMessage = NormalizeApplicationEventMessage(message);
                EventLog.WriteEntry(EventSourceName, normalizedMessage, EventLogEntryType.Information, eventId);
                WriteLog("write event", $"application event written successfully; eventId={eventId}");
            }
            catch (Exception ex)
            {
                WriteLog("write event", $"failed to write application event{Environment.NewLine}{ex}");
            }
        }

        private static string NormalizeApplicationEventMessage(string message)
        {
            string normalized = message ?? string.Empty;

            if (normalized.Length <= EventLogMessageMaxLength)
            {
                return normalized;
            }

            return normalized.Substring(0, EventLogMessageMaxLength)
                + Environment.NewLine
                + "[message truncated before writing to Windows Event Log]";
        }

        private static bool IsApplicationLogWritable(out string failureReason)
        {
            failureReason = string.Empty;

            try
            {
                if (!EventLog.Exists(ApplicationLogName))
                {
                    failureReason = $"Event log '{ApplicationLogName}' does not exist.";
                    return false;
                }

                if (!EventLog.SourceExists(EventSourceName))
                {
                    EventLog.CreateEventSource(new EventSourceCreationData(EventSourceName, ApplicationLogName));
                    failureReason = $"Event source '{EventSourceName}' created. Restart the application to enable event logging.";
                    return false;
                }

                using (var eventLog = new EventLog(ApplicationLogName))
                {
                    eventLog.Source = EventSourceName;
                }

                return true;
            }
            catch (Exception ex)
            {
                failureReason = ex.Message;
                return false;
            }
        }

        private sealed class SCCMInfoServiceHost : ServiceBase
        {
            protected override void OnStart(string[] args)
            {
                IsServiceMode = true;
                WriteLog("service", "SCCMInfo service started");
                CCMMon();
            }

            protected override void OnStop()
            {
                WriteLog("service", "SCCMInfo service stopping");
                StopMonitoring();
            }

            protected override void OnShutdown()
            {
                OnStop();
                base.OnShutdown();
            }
        }
    }
}
