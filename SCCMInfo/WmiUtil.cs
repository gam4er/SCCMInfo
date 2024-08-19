using System;
using System.Management;
namespace SCCMInfo
{
    public class WmiUtil
    {
        public static ManagementScope NewWmiConnection()
        {
            string path = "", server = "", siteCode = "";
            ConnectionOptions connection = new ConnectionOptions();
            (server, siteCode) = GetCurrentManagementPointAndSiteCode();
            path = $"\\\\{server}\\root\\SMS\\site_{siteCode}";

            ManagementScope wmiConnection = null;
            try
            {
                if (!string.IsNullOrEmpty(path))
                {
                    wmiConnection = new ManagementScope(path, connection);
                    Console.WriteLine($"[+] Connecting to {wmiConnection.Path}");
                    wmiConnection.Connect();
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"[!] Access to the WMI provider was not authorized: {ex.Message.Trim()}");
            }
            catch (ManagementException ex)
            {
                Console.WriteLine($"[!] Could not connect to {path}: {ex.Message}");
                if (path.Contains("\\root\\CCM") && ex.Message == "Invalid namespace ")
                {
                    Console.WriteLine(
                        "[!] The SCCM client may not be installed on this machine\n" +
                        "[!] Try specifying an SMS Provider and site code"
                    );
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unhandled exception of type {ex.GetType()} occurred: {ex.Message}");
            }

            return wmiConnection;
        }

        private static (string, string) GetCurrentManagementPointAndSiteCode()
        {
            string siteCode = null;
            string managementPoint = null;

            try
            {
                // Подключение к пространству имен root\ccm
                ManagementScope scope = new ManagementScope(@"\\.\ROOT\ccm");
                scope.Connect();

                // Запрос для получения CurrentManagementPoint и Site Code (Name)
                ObjectQuery query = new ObjectQuery("SELECT CurrentManagementPoint, Name FROM SMS_Authority");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
                ManagementObjectCollection results = searcher.Get();

                foreach (ManagementObject result in results)
                {
                    managementPoint = result ["CurrentManagementPoint"]?.ToString();
                    siteCode = result ["Name"]?.ToString().Replace("SMS:","");
                    break; // Обычно интересует первый найденный результат
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving ManagementPoint and SiteCode: {ex.Message}");
            }

            Console.WriteLine($"MP: {managementPoint} site: {siteCode}");
            return (managementPoint, siteCode);
        }
    }

}
