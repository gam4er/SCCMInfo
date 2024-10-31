using Spectre.Console;

using System;
using System.IO;
using System.Management;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace SCCMInfo
{
    internal class SCCMInfo
    {
        public static ManagementScope scope = new ManagementScope();

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

                // Если класс объекта SMS_DeploymentInfo, получаем связанные данные
                if (targetInstance.ClassPath.ClassName == "SMS_DeploymentInfo")
                {
                    string collectionID = targetInstance ["CollectionID"]?.ToString();
                    string targetName = targetInstance ["TargetName"]?.ToString();
                    string deploymentName = targetInstance ["DeploymentName"]?.ToString();
                    string collectionName = targetInstance ["CollectionName"]?.ToString();

                    // Добавляем информацию в таблицу
                    table.AddRow("CollectionID", collectionID);
                    table.AddRow("CollectionName", collectionName);

                    // Получаем директорию для файлов
                    //string logDirectory = "c:\\temp";


                    string queryString = $"SELECT * FROM SMS_FullCollectionMembership WHERE CollectionID='{collectionID}'";
                    ObjectQuery query = new ObjectQuery(queryString);
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

                    ManagementObjectCollection collectionMembers = searcher.Get();

                    // Создаем CSV данные
                    var csvData = new StringBuilder();
                    csvData.AppendLine("Name,ResourceID,ResourceType,ResourceDomainORWorkgroup,LastChangeTime,IsDirect");

                    foreach (ManagementObject member in collectionMembers)
                    {
                        member.Get();
                        string name = "";
                        string resourceID = "";
                        string resourceType = "";
                        string domain = "";
                        string isClient = "";
                        string isActive = "";
                        string isApproved = "";
                        string isBlocked = "";

                        try { name = member ["Name"]?.ToString() ?? ""; }
                        catch { /* игнорируем */ }

                        try { resourceID = member ["ResourceID"]?.ToString() ?? ""; }
                        catch { /* игнорируем */ }

                        try { resourceType = member ["ResourceType"]?.ToString() ?? ""; }
                        catch { /* игнорируем */ }

                        try { domain = member ["Domain"]?.ToString() ?? ""; }
                        catch { /* игнорируем */ }

                        try { isClient = member ["IsClient"]?.ToString() ?? ""; }
                        catch { /* игнорируем */ }

                        try { isActive = member ["IsActive"]?.ToString() ?? ""; }
                        catch { /* игнорируем */ }

                        try { isApproved = member ["IsApproved"]?.ToString() ?? ""; }
                        catch { /* игнорируем */ }

                        try { isBlocked = member ["IsBlocked"]?.ToString() ?? ""; }
                        catch { /* игнорируем */ }

                        csvData.AppendLine($"{name},{resourceID},{resourceType},{domain},{isClient},{isActive},{isApproved},{isBlocked}");
                    }

                    // Сохраняем CSV файл
                    string csvFileName = $"{deploymentName}_{collectionID}.csv";                    
                    File.WriteAllText(csvFileName, csvData.ToString());

                    int membersCount = collectionMembers.Count;
                    table.AddRow("CollectionMembersCount", membersCount.ToString());
                    table.AddRow("CollectionMembersFile", csvFileName);


                    logMessage.AppendLine($"Collection members info saved to file {csvFileName}, total members: {membersCount}");
                    logMessage.AppendLine($"Collection Name: {collectionName}, ID: {collectionID}");

                    // Get Application Information                    
                    string appQueryStr = $"SELECT * FROM SMS_ApplicationLatest WHERE LocalizedDisplayName LIKE '{targetName}'";

                    ObjectQuery appQuery = new ObjectQuery(appQueryStr);
                    ManagementObjectSearcher appSearcher = new ManagementObjectSearcher(scope, appQuery);
                    ManagementObjectCollection applications = appSearcher.Get();

                    ManagementObject application = null;
                    foreach (ManagementObject app_in_collection in applications)
                    {
                        try
                        {
                            application = app_in_collection;
                            application.Get();
                            logMessage.AppendLine("App refreshed");
                            // Get SDMPackageXML property
                            string sdmPackageXml = application ["SDMPackageXML"]?.ToString();                             
                            logMessage.AppendLine($"sdmPackageXml gotted");
                            try
                            {
                                XmlDocument xmlDoc = new XmlDocument();
                                xmlDoc.LoadXml(sdmPackageXml);

                                // Sanitize file name
                                string safeName = Regex.Replace(targetName, @"[\\/:*?""<>|\r\n]+", "_");
                                safeName = Regex.Replace(safeName, @"\s+", "_");
                                string xmlFileName = $"{deploymentName}_{safeName}.xml";

                                // Save XML to file
                                xmlDoc.Save(xmlFileName);
                                logMessage.AppendLine($"Application info saved to file {xmlFileName}");

                                // Extract information from XML
                                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                                nsmgr.AddNamespace("ab", "http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest");

                                XmlNode appMgmtDigest = xmlDoc.SelectSingleNode("//ab:AppMgmtDigest", nsmgr);
                                if (appMgmtDigest != null)
                                {

                                    string appName = appMgmtDigest.SelectSingleNode("ab:Name", nsmgr)?.InnerText ?? "";
                                    string appVersion = appMgmtDigest.SelectSingleNode("ab:Version", nsmgr)?.InnerText ?? "";
                                    string appPublisher = appMgmtDigest.SelectSingleNode("ab:Publisher", nsmgr)?.InnerText ?? "";
                                    string appInstallDate = appMgmtDigest.SelectSingleNode("ab:InstallDate", nsmgr)?.InnerText ?? "";
                                    string appInstallSource = appMgmtDigest.SelectSingleNode("ab:InstallSource", nsmgr)?.InnerText ?? "";
                                                                        
                                    // Добавляем информацию в таблицу
                                    table.AddRow("ApplicationName", appName);
                                    table.AddRow("ApplicationVersion", appVersion);
                                    table.AddRow("ApplicationPublisher", appPublisher);
                                    table.AddRow("ApplicationInstallDate", appInstallDate);
                                    table.AddRow("ApplicationInstallSource", appInstallSource);
                                    // Добавляем путь к файлу в таблицу
                                    table.AddRow("ApplicationXMLFile", xmlFileName);
                                    logMessage.AppendLine($"Application info saved to vars");

                                    try
                                    {
                                        // Get Install Arguments
                                        XmlNodeList argNodes = xmlDoc.SelectNodes("//ab:DeploymentType/ab:Installer/ab:InstallAction/ab:Args/ab:Arg", nsmgr);
                                        foreach (XmlNode argNode in argNodes)
                                        {
                                            string argName = argNode.Attributes ["Name"]?.Value;
                                            string argText = argNode.InnerText;
                                            logMessage.AppendLine($"And that application install arguments:       '{argName}' '{argText}'");
                                            table.AddRow(argName, argText);
                                        }
                                        logMessage.AppendLine($"Application Install Arguments saved");
                                    } 
                                    catch {
                                        logMessage.AppendLine("Error while processing properties SMS_ApplicationLatest: Install Arguments not found");
                                    }

                                    try
                                    {
                                        // Get CommandLineArg
                                        XmlNode commandLineArgNode = xmlDoc.SelectSingleNode("//ab:DeploymentType/ab:Installer/ab:InstallAction/ab:Args/ab:Arg[@Name='CommandLine']", nsmgr);
                                        string commandLineArg = commandLineArgNode?.InnerText;
                                        table.AddRow("application command line", commandLineArg);
                                        logMessage.AppendLine($"And that application command line:             '{commandLineArg}'");
                                    }
                                    catch
                                    {
                                        logMessage.AppendLine("Error while processing properties SMS_ApplicationLatest: CommandLineArg not found");
                                    }

                                    try
                                    {
                                        // Get InstallCommandLine
                                        string installCommandLine = xmlDoc.SelectSingleNode("//ab:DeploymentType/ab:Installer/ab:CustomData/ab:InstallCommandLine", nsmgr)?.InnerText;
                                        logMessage.AppendLine($"And that application install command line:     '{installCommandLine}'");
                                        table.AddRow("application install command line", installCommandLine);
                                    }
                                    catch
                                    {
                                        logMessage.AppendLine("Error while processing properties SMS_ApplicationLatest: InstallCommandLine not found");
                                    }

                                    try
                                    {
                                        // Get UninstallCommandLine
                                        string uninstallCommandLine = xmlDoc.SelectSingleNode("//ab:DeploymentType/ab:Installer/ab:CustomData/ab:UninstallCommandLine", nsmgr)?.InnerText;
                                        logMessage.AppendLine($"And that application uninstall command line:   '{uninstallCommandLine}'");
                                        table.AddRow("application uninstall command line", uninstallCommandLine);
                                    }
                                    catch
                                    {
                                        logMessage.AppendLine("Error while processing properties SMS_ApplicationLatest: UninstallCommandLine not found");
                                    }
                                }
                                else
                                {
                                    logMessage.AppendLine("Error while processing properties SMS_ApplicationLatest: AppMgmtDigest not found");
                                }
                            }
                            catch (Exception ex)
                            {
                                logMessage.AppendLine($"Error while processing properties SMS_ApplicationLatest: {ex.Message}");
                            }
                        }
                        catch (Exception ex)
                        {
                            logMessage.AppendLine($"Error while getting properties SMS_ApplicationLatest: {ex.Message}");
                        }
                        break;
                    }                    
                }

                // Записываем в лог
                WriteLog(logMessage.ToString());
                // Выводим таблицу на экран
                AnsiConsole.Write(table);
            }
            catch (Exception ex)
            {
                // Логируем ошибку
                WriteLog($"HandleEvent error: {ex.Message}\n{ex.StackTrace}");
            }
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
