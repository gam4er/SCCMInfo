using System;
using System.Collections;
using System.EnterpriseServices.Internal;
using System.IO;
using System.Management;
using System.Management.Instrumentation;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

[assembly: WmiConfiguration(@"root\cimv2", HostingModel = ManagementHostingModel.Decoupled)]
namespace WMIEventProvider
{
    [System.ComponentModel.RunInstaller(true)]
    public class InstallCCMMonitoring : DefaultManagementInstaller
    {
        public override void Install(IDictionary stateSaver)
        {

            Publish publish = new Publish();
            publish.GacInstall("DeploymentInfoLogger.dll");
            base.Install(stateSaver);
            RegistrationServices RS = new RegistrationServices();

            //This should be fixed with .NET 3.5 SP1
            RS.RegisterAssembly(System.Reflection.Assembly.GetExecutingAssembly(), AssemblyRegistrationFlags.SetCodeBase);
            //InstrumentationManager.RegisterType(typeof(NewProcessInfoLogger));
            //NewProcessInfoLogger.Start();            
            var t = new NewProcessInfoLogger();
        }

        public override void Uninstall(IDictionary savedState)
        {
            try
            {
                Publish publish = new Publish();
                publish.GacRemove("DeploymentInfoLogger.dll");
            }
            catch { }

            try
            {
                base.Uninstall(savedState);
            }
            catch { }
        }
    }

    [ManagementEntity(External = true,Singleton = true)]
    
    public class NewProcessInfoLogger 
    {
        //[ManagementKey]
        private static NewProcessInfoLogger _instance = new NewProcessInfoLogger(); 

        private static ManagementEventWatcher _watcher;
        
        [ManagementKey]
        public string Member { get; set; }

        //[ManagementBind]
        [ManagementCreate]
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
