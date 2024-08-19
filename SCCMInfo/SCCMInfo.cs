using System;
using System.IO;
using System.Management;
using System.Text;

namespace SCCMInfo
{
    internal class SCCMInfo
    {

        private static void ProcMon()
        {
            ManagementEventWatcher _watcher;
            WriteLog("Starting ProcessInfoLogger");
            try
            {
                WqlEventQuery query = new WqlEventQuery(
                    "SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance isa \"Win32_Process\""
                );
                _watcher = new ManagementEventWatcher(query);
                _watcher.EventArrived += new EventArrivedEventHandler(HandleEvent);
                _watcher.Start();

            }
            catch (Exception ex)
            {
            }
        }

        private static void CCMMon()
        {
            ManagementEventWatcher _watcher;
            WriteLog("Starting ProcessInfoLogger");
            try
            {
                WqlEventQuery query = new WqlEventQuery(
                    "__InstanceCreationEvent",
                    new TimeSpan(0, 0, 1),
                    "TargetInstance ISA 'SMS_DeploymentInfo'");

                ManagementScope scope = WmiUtil.NewWmiConnection();
                WriteLog(scope.Path.Path);
                _watcher = new ManagementEventWatcher(scope,query);
                _watcher.EventArrived += new EventArrivedEventHandler(HandleEvent);
                _watcher.Start();

            }
            catch (Exception ex)
            {
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

                // Проходим по всем свойствам объекта
                foreach (PropertyData property in targetInstance.Properties)
                {
                    string propertyName = property.Name;
                    string propertyValue = property.Value != null ? property.Value.ToString() : "null";

                    // Форматируем строку с табуляцией для выравнивания
                    logMessage.AppendLine($"{propertyName.PadRight(30)}\t{propertyValue}");
                }

                logMessage.AppendLine();

                // Записываем в лог
                WriteLog(logMessage.ToString());
            }
            catch (Exception ex)
            {
                // Логируем ошибку
                WriteLog($"HandleEvent error: {ex.Message}\n{ex.StackTrace}");
            }
        }

        public static void WriteLog(string message)
        {
            string logFilePath = "c:\\temp\\ProcessInfoLog.txt";
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
