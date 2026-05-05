using System;
using System.Collections;
// using System.EnterpriseServices.Internal; // Removed: Not available in .NET 8.0 and not present in project dependencies
using System.IO;
using System.Management;
// using System.Management.Instrumentation; // Removed: Not available in .NET 8.0 or current dependencies
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

// [assembly: WmiConfiguration(@"root\cimv2", HostingModel = ManagementHostingModel.Decoupled)] // Removed: WmiConfigurationAttribute and ManagementHostingModel are not available in .NET 8.0 or current dependencies
namespace WMIEventProvider
{
    [System.ComponentModel.RunInstaller(true)]
    // DefaultManagementInstaller is not available in .NET 8.0 or current dependencies.
    // Consider implementing custom installer logic or using a different base class if required.
    public class InstallCCMMonitoring
    {
        public void Install(IDictionary stateSaver)
        {

            // The 'Publish' type is not available in .NET 8.0 or your current dependencies. GAC operations are not supported directly. Consider alternative deployment strategies or use a compatible library if GAC install is required.
            // publish.GacInstall("DeploymentInfoLogger.dll"); // Not supported in .NET 8.0
            // base.Install(stateSaver); // No base class to call, so this is commented out.
            // RegistrationServices RS = new RegistrationServices(); // Not available in .NET 8.0 and not present in project dependencies

            //This should be fixed with .NET 3.5 SP1
            // RS.RegisterAssembly(System.Reflection.Assembly.GetExecutingAssembly(), AssemblyRegistrationFlags.SetCodeBase); // Not available in .NET 8.0 and not present in project dependencies
            //InstrumentationManager.RegisterType(typeof(NewProcessInfoLogger));
            //NewProcessInfoLogger.Start();            
            var t = new NewProcessInfoLogger();
        }

        public void Uninstall(IDictionary savedState)
        {
            try
            {
                // The 'Publish' type is not available in .NET 8.0 or your current dependencies. GAC operations are not supported directly. Consider alternative deployment strategies or use a compatible library if GAC remove is required.
                // publish.GacRemove("DeploymentInfoLogger.dll"); // Not supported in .NET 8.0
            }
            catch { }

            try
            {
                // base.Uninstall(savedState); // No base class to call, so this is commented out.
            }
            catch { }
        }
    }

    // [ManagementEntity(External = true,Singleton = true)] // ManagementEntityAttribute is not available in .NET 8.0 or current dependencies
    
    public class NewProcessInfoLogger 
    {
        //[ManagementKey]
        private static NewProcessInfoLogger _instance = new NewProcessInfoLogger(); 

        private static ManagementEventWatcher _watcher;
        
        // [ManagementKey] // ManagementKeyAttribute is not available in .NET 8.0 or current dependencies
        public string Member { get; set; }

        //[ManagementBind]
        // [ManagementCreate] // ManagementCreateAttribute is not available in .NET 8.0 or current dependencies
        static NewProcessInfoLogger()
        {
            _instance = new NewProcessInfoLogger();
            _instance.Start();    
            //Thread.Sleep(Timeout.Infinite);
        }

        //[ManagementTask]
        //[ManagementBind]
        public void Start()
        {
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
                WriteLog($"Stop error: {ex.Message}\n{ex.StackTrace}");
            }
        }
        public static void Stop()
        {
            if (_watcher != null)
            {
                _watcher.Stop();
                _watcher.Dispose();
                //_instance = null;
            }
        }
        private static void HandleEvent(object sender, EventArrivedEventArgs e)
        {
            try
            {
                ManagementBaseObject targetInstance = (ManagementBaseObject)e.NewEvent ["TargetInstance"];
                string Name = targetInstance ["Name"]?.ToString();
                string ExecutablePath = targetInstance ["ExecutablePath"]?.ToString();

                string logMessage = $"Name: {Name}, ExecutablePath: {ExecutablePath}";
                WriteLog(logMessage);
            }
            catch (Exception ex)
            {
                WriteLog($"HandleEvent error: {ex.Message}\n{ex.StackTrace}");
            }

        }
        private static void WriteLog(string message)
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
                    writer.WriteLine($"{DateTime.Now}: {message}");
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
