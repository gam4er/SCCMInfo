using System;
using System;
using System.Management;

namespace SCCMInfo
{
    public class WmiUtil
    {
        public static ManagementScope NewWmiConnection()
        {
            string server = string.Empty;
            string siteCode = string.Empty;
            ConnectionOptions connection = new ConnectionOptions();
            (server, siteCode) = GetCurrentManagementPointAndSiteCode();

            if (string.IsNullOrWhiteSpace(server) || string.IsNullOrWhiteSpace(siteCode))
            {
                global::SCCMInfo.SCCMInfo.WriteLog(
                    "wmi connection",
                    $"connection aborted because management point or site code is empty; managementPoint={server}; siteCode={siteCode}");
                return null;
            }

            string path = $"\\\\{server}\\root\\SMS\\site_{siteCode}";
            ManagementScope wmiConnection = null;

            try
            {
                wmiConnection = new ManagementScope(path, connection);
                global::SCCMInfo.SCCMInfo.WriteLog("wmi connection", $"connecting to {wmiConnection.Path}");
                wmiConnection.Connect();
                global::SCCMInfo.SCCMInfo.WriteLog("wmi connection", $"connection established to {wmiConnection.Path}");
            }
            catch (UnauthorizedAccessException ex)
            {
                global::SCCMInfo.SCCMInfo.WriteLog("wmi connection", $"access denied: {ex.Message.Trim()}");
            }
            catch (ManagementException ex)
            {
                global::SCCMInfo.SCCMInfo.WriteLog("wmi connection", $"connection failed for path {path}{Environment.NewLine}{ex}");
                if (path.Contains("\\root\\CCM") && ex.Message == "Invalid namespace ")
                {
                    global::SCCMInfo.SCCMInfo.WriteLog(
                        "wmi connection",
                        "the SCCM client may not be installed on this machine; try specifying an SMS Provider and site code");
                }
            }
            catch (Exception ex)
            {
                global::SCCMInfo.SCCMInfo.WriteLog("wmi connection", $"unexpected connection error{Environment.NewLine}{ex}");
            }

            return wmiConnection;
        }

        private static (string, string) GetCurrentManagementPointAndSiteCode()
        {
            string siteCode = null;
            string managementPoint = null;

            try
            {
                ManagementScope scope = new ManagementScope(@"\\.\ROOT\ccm");
                scope.Connect();

                ObjectQuery query = new ObjectQuery("SELECT CurrentManagementPoint, Name FROM SMS_Authority");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);
                ManagementObjectCollection results = searcher.Get();

                foreach (ManagementObject result in results)
                {
                    managementPoint = result["CurrentManagementPoint"]?.ToString();
                    siteCode = result["Name"]?.ToString().Replace("SMS:", "");
                    break;
                }
            }
            catch (Exception ex)
            {
                global::SCCMInfo.SCCMInfo.WriteLog("wmi connection", $"failed to retrieve management point and site code{Environment.NewLine}{ex}");
            }

            global::SCCMInfo.SCCMInfo.WriteLog("wmi connection", $"management point={managementPoint}; siteCode={siteCode}");
            return (managementPoint, siteCode);
        }
    }
}
