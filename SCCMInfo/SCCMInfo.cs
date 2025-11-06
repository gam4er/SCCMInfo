using Spectre.Console;

using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Text;

using SCCMInfo.Enrichers;

namespace SCCMInfo
{
    internal class SCCMInfo
    {
        public static ManagementScope scope = new ManagementScope();

        private const string WindowsEventLogName = "Application";
        private const string WindowsEventSource = "SCCMInfoWatcher";
        private const int WindowsEventId = 1024;

        private static readonly IInstanceEnricher [] InstanceEnrichers =
        {
            new SmsDeploymentInfoEnricher(),
            new SmsCombinedDeviceResourcesEnricher(),
            new SmsAdminEnricher(),
            new SmsScriptsEnricher(),
            new SmsSciReservedEnricher()
        };

        private static void ProcMon()
        {
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
                }

            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }

        private static void CCMMon()
        {
            
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
                /*
                ConnectionOptions options = new ConnectionOptions();
                options.Username = @"gam\administrator";
                options.Password = @"~=Gam1987=~";                
                scope = new ManagementScope(@"\\192.168.199.129\ROOT\SMS\site_GAM", options);
                scope.Connect();
                */
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
                }

            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }

        static void Main(string [] args)
        {
            //ProcMon();
            CCMMon();
            Console.ReadLine();
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

                // Записываем в лог
                WriteLog(logMessage.ToString());

                // Пытаемся записать таблицу в журнал приложений Windows
                TryWriteTableToWindowsEventLog(table);

                // Выводим таблицу на экран
                AnsiConsole.Write(table);
            }
            catch (Exception ex)
            {
                // Логируем ошибку
                WriteLog($"HandleEvent error: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private static void TryWriteTableToWindowsEventLog(Table table)
        {
            if (Environment.OSVersion.Platform != PlatformID.Win32NT)
            {
                return;
            }

            try
            {
                if (!EventLog.Exists(WindowsEventLogName))
                {
                    WriteLog($"Windows event log '{WindowsEventLogName}' not found. Skipping event log write.");
                    return;
                }

                EnsureEventSourceExists();

                if (!EventLog.SourceExists(WindowsEventSource) ||
                    !string.Equals(EventLog.LogNameFromSourceName(WindowsEventSource, "."), WindowsEventLogName, StringComparison.OrdinalIgnoreCase))
                {
                    WriteLog($"Windows event source '{WindowsEventSource}' is unavailable for log '{WindowsEventLogName}'. Skipping event log write.");
                    return;
                }

                string tableContent = RenderTableToText(table);

                if (!string.IsNullOrWhiteSpace(tableContent))
                {
                    EventLog.WriteEntry(WindowsEventSource, tableContent, EventLogEntryType.Information, WindowsEventId);
                }
            }
            catch (Exception ex)
            {
                WriteLog($"Failed to write table to Windows Event Log: {ex.Message}");
            }
        }

        private static void EnsureEventSourceExists()
        {
            if (EventLog.SourceExists(WindowsEventSource))
            {
                return;
            }

            try
            {
                EventSourceCreationData sourceData = new EventSourceCreationData(WindowsEventSource, WindowsEventLogName);
                EventLog.CreateEventSource(sourceData);
            }
            catch (Exception ex)
            {
                WriteLog($"Unable to create Windows event source '{WindowsEventSource}': {ex.Message}");
            }
        }

        private static string RenderTableToText(Table table)
        {
            var recorder = AnsiConsole.Record();
            recorder.Write(table);
            return recorder.ExportText();
        }

        public static void WriteLog(string message)
        {
            string logFilePath = @".\ProcessInfoLog.txt";
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
                Console.WriteLine($"Log write error: {ex.Message}\n{ex.StackTrace}");
            }
        }

    }
}
