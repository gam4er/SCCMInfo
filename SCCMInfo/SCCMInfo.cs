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
        internal const int SmsCombinedDeviceResourcesEventId = 2002;

        private static readonly IInstanceEnricher [] InstanceEnrichers =
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
            WriteLog("Starting ProcessInfoLogger");

            string [] classesToMonitor = {
                "Win32_Process"
            };

            try
            {
                ManagementScope scope = new ManagementScope();
                WriteLog(scope.Path.Path);

                foreach (string className in classesToMonitor)
                {
                    WqlEventQuery query = new WqlEventQuery(
                        "__InstanceCreationEvent",
                        new TimeSpan(0, 0, 1),
                        $"TargetInstance ISA '{className}'");

                    ManagementEventWatcher _watcher;
                    _watcher = new ManagementEventWatcher(scope, query);
                    _watcher.EventArrived += new EventArrivedEventHandler(HandleEvent);
                    _watcher.Start();
                    ActiveWatchers.Add(_watcher);
                }

            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }

        private static void CCMMon()
        {
            StopMonitoring();

            WriteLog("Starting ProcessInfoLogger");
            
            string [] classesToMonitor = {
                "SMS_DeploymentInfo",
                "SMS_CombinedDeviceResources",
                "SMS_Admin",
                "SMS_Scripts",
                "SMS_SCI_Reserved"
            };

            try
            {
                scope = WmiUtil.NewWmiConnection();
                WriteLog(scope.Path.Path);

                foreach (string className in classesToMonitor)
                {
                    WqlEventQuery query = new WqlEventQuery(
                        "__InstanceCreationEvent",
                        new TimeSpan(0, 0, 1),
                        $"TargetInstance ISA '{className}'");

                    ManagementEventWatcher _watcher;
                    _watcher = new ManagementEventWatcher(scope, query);
                    _watcher.EventArrived += new EventArrivedEventHandler(HandleEvent);
                    _watcher.Start();
                    ActiveWatchers.Add(_watcher);
                }

            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }

        static void Main(string [] args)
        {
            args = args ?? Array.Empty<string>();

            if (args.Any(a => string.Equals(a, "--install", StringComparison.OrdinalIgnoreCase)))
            {
                InstallService();
                return;
            }

            bool runAsService = args.Any(a => string.Equals(a, "--service", StringComparison.OrdinalIgnoreCase));
            IsServiceMode = runAsService || !Environment.UserInteractive;

            if (IsServiceMode)
            {
                ServiceBase.Run(new SCCMInfoServiceHost());
                return;
            }

            Console.CancelKeyPress += (sender, eventArgs) =>
            {
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
                return;
            }

            try
            {
                if (IsServiceInstalled(ServiceName))
                {
                    Console.WriteLine($"Сервис \"{ServiceName}\" уже установлен.");
                    return;
                }

                string executablePath = Process.GetCurrentProcess().MainModule.FileName;
                string arguments = $"create \"{ServiceName}\" binPath= \"\\\"{executablePath}\\\" --service\" start= auto DisplayName= \"{ServiceDisplayName}\"";

                ProcessStartInfo startInfo = new ProcessStartInfo("sc.exe", arguments)
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
                        return;
                    }

                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode == 0)
                    {
                        Console.WriteLine("Сервис успешно установлен.");
                        if (!string.IsNullOrWhiteSpace(output))
                        {
                            Console.WriteLine(output.Trim());
                        }

                        WriteLog("Service installed successfully.");
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
                WriteLog($"Service installation failed: {ex.Message}");
            }
        }

        private static bool IsServiceInstalled(string serviceName)
        {
            try
            {
                return ServiceController.GetServices().Any(service => string.Equals(service.ServiceName, serviceName, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                WriteLog($"Failed to determine whether service '{serviceName}' is installed: {ex.Message}");
                return false;
            }
        }

        private static bool IsRunningOnWindows()
        {
            PlatformID platform = Environment.OSVersion.Platform;
            return platform == PlatformID.Win32NT || platform == PlatformID.Win32S || platform == PlatformID.Win32Windows || platform == PlatformID.WinCE;
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
                }
                catch (Exception ex)
                {
                    WriteLog($"Failed to stop watcher: {ex.Message}");
                }
                finally
                {
                    ActiveWatchers.Remove(watcher);
                }
            }
        }

        private static void HandleEvent(object sender, EventArrivedEventArgs e)
        {
            try
            {
                // Получаем объект, который был создан (TargetInstance)
                ManagementBaseObject targetInstance = (ManagementBaseObject)e.NewEvent ["TargetInstance"];

                // Инициализируем строку для логирования
                StringBuilder logMessage = new StringBuilder();

                // Получаем текущее время
                string creationTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // Заголовок для класса объекта и времени создания
                logMessage.AppendLine($"Instance created: {targetInstance.ClassPath.ClassName} at {creationTime}");

                // Инициализируем таблицу для вывода на экран
                var table = new Table();
                table.Title($"[bold yellow]Instance created: {targetInstance.ClassPath.ClassName}[/]");
                table.AddColumn("Property");
                table.AddColumn("Value");

                // Проходим по всем свойствам объекта
                foreach (PropertyData property in targetInstance.Properties)
                {
                    string propertyName = property.Name;
                    string propertyValue = property.Value != null ? property.Value.ToString() : "null";

                    // Добавляем строку в лог
                    logMessage.AppendLine($"{propertyName.PadRight(30)}\t{propertyValue}");

                    // Добавляем строку в таблицу
                    table.AddRow(propertyName, propertyValue);
                }

                logMessage.AppendLine();

                foreach (IInstanceEnricher enricher in InstanceEnrichers)
                {
                    if (string.Equals(enricher.ClassName, targetInstance.ClassPath.ClassName, StringComparison.OrdinalIgnoreCase))
                    {
                        enricher.Enrich(targetInstance, table, logMessage, scope);
                        break;
                    }
                }

                string logText = logMessage.ToString();

                // Записываем в лог
                WriteLog(logText);

                // Пишем в журнал приложений
                WriteApplicationEvent(logText);
                // Выводим таблицу на экран
                if (!IsServiceMode)
                {
                    AnsiConsole.Write(table);
                }
            }
            catch (Exception ex)
            {
                // Логируем ошибку
                WriteLog($"HandleEvent error: {ex.Message}\n{ex.StackTrace}");
            }
        }

        public static void WriteLog(string message)
        {
            string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ProcessInfoLog.txt");
            try
            {
                // Получаем директорию из пути
                string logDirectory = Path.GetDirectoryName(logFilePath);

                // Проверяем, существует ли директория, и создаем ее, если не существует
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                // Создаем или открываем файл и записываем в него лог
                using (StreamWriter writer = new StreamWriter(logFilePath, true, Encoding.UTF8))
                {
                    writer.WriteLine(message);
                }
            }
            catch (Exception ex)
            {
                // Обработка исключений, если требуется
                if (!IsServiceMode)
                {
                    Console.WriteLine($"Log write error: {ex.Message}\n{ex.StackTrace}");
                }
            }
        }

        internal static void WriteApplicationEvent(string message, int eventId = EventLogEntryId)
        {
            if (!IsApplicationLogWritable(out string failureReason))
            {
                if (!string.IsNullOrWhiteSpace(failureReason))
                {
                    WriteLog($"Application event log is not writable: {failureReason}");
                }
                return;
            }

            try
            {
                EventLog.WriteEntry(EventSourceName, message, EventLogEntryType.Information, eventId);
            }
            catch (Exception ex)
            {
                WriteLog($"Failed to write to application event log: {ex.Message}\n{ex.StackTrace}");
            }
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

                using (EventLog eventLog = new EventLog(ApplicationLogName))
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
            protected override void OnStart(string [] args)
            {
                IsServiceMode = true;
                WriteLog("SCCMInfo service started.");
                CCMMon();
            }

            protected override void OnStop()
            {
                WriteLog("SCCMInfo service stopping.");
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
