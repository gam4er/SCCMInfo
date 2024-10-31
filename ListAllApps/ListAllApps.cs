using System;
using System.IO;
using System.Management;
using System.Text.RegularExpressions;
using System.Xml;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Security.AccessControl;

namespace ListAllApps
{
    class ListAllApps
    {
        static void Main(string [] args)
        {
            ConnectionOptions options = new ConnectionOptions();
            options.Username = @"";
            options.Password = "";
            ManagementScope deploymentScope = new ManagementScope(@"\\192.168.199.129\ROOT\SMS\site_GAM", options);
            deploymentScope.Connect();

            ObjectQuery deploymentQuery = new ObjectQuery("SELECT * FROM SMS_DeploymentInfo");
            ManagementObjectSearcher deploymentSearcher = new ManagementObjectSearcher(deploymentScope, deploymentQuery);
            ManagementObjectCollection deployments = deploymentSearcher.Get();

            Console.WriteLine($"WARNING: Finished getting all deployments from ROOT\\SMS\\site_GAM");

            foreach (ManagementObject deployment in deployments)
            {
                string targetName = deployment ["TargetName"]?.ToString();
                string deploymentName = deployment ["DeploymentName"]?.ToString();
                Console.WriteLine($"WARNING: Find Deployment '{targetName}' with DeploymentName '{deploymentName}'");

                string collectionID = deployment ["CollectionID"]?.ToString();
                string collectionName = deployment ["CollectionName"]?.ToString();
                Console.WriteLine($"WARNING:    With that deployment associated collection '{collectionName}' with ID '{collectionID}'");

                // Get Collection Members
                string collectionQueryStr = $"SELECT * FROM SMS_FullCollectionMembership WHERE CollectionID='{collectionID}'";
                ObjectQuery collectionQuery = new ObjectQuery(collectionQueryStr);
                ManagementObjectSearcher collectionSearcher = new ManagementObjectSearcher(deploymentScope, collectionQuery);
                ManagementObjectCollection collectionMembers = collectionSearcher.Get();

                // Create a list of objects for export
                var csvData = new List<dynamic>();
                foreach (ManagementObject member in collectionMembers)
                {
                    member.Get();
                    try
                    {
                        var obj = new
                        {
                            name = member ["Name"]?.ToString() ?? "",
                            resourceID = member ["ResourceID"]?.ToString() ?? "",
                            resourceType = member ["ResourceType"]?.ToString() ?? "",
                            domain = member ["Domain"]?.ToString() ?? "",
                            isClient = member ["IsClient"]?.ToString() ?? "",
                            isActive = member ["IsActive"]?.ToString() ?? "",
                            isApproved = member ["IsApproved"]?.ToString() ?? "",
                            isBlocked = member ["IsBlocked"]?.ToString() ?? "",
                        };
                        csvData.Add(obj);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"WARNING: Error getting collection member: {ex.Message}");
                        continue;
                    }
                    
                }

                // Export data to CSV file
                string csvFileName = $"{deploymentName}_{collectionID}.csv";
                using (var writer = new StreamWriter(csvFileName))
                {
                    // Write header
                    writer.WriteLine("Name,ResourceID,ResourceType,ResourceDomainORWorkgroup,LastChangeTime,IsDirect");

                    // Write data
                    foreach (var item in csvData)
                    {
                        writer.WriteLine($"{item.name}," +
                            $"{item.resourceID}," +
                            $"{item.resourceType}," +
                            $"{item.domain}," +
                            $"{item.isClient}," +
                            $"{item.isActive}," +
                            $"{item.isApproved}," +
                            $"{item.isBlocked}");
                    }
                }

                Console.WriteLine($"Information about collection members {collectionName} with ID {collectionID}: {collectionMembers.Count} items were written to file {csvFileName}");

                string localizedDisplayName = deployment ["TargetName"]?.ToString();
                Console.WriteLine($"WARNING:    And that deployment associated application '{localizedDisplayName}'");

                // Get Application Information
                string appQueryStr = $"SELECT * FROM SMS_ApplicationLatest WHERE LocalizedDisplayName LIKE '{localizedDisplayName}'";
                ObjectQuery appQuery = new ObjectQuery(appQueryStr);
                ManagementObjectSearcher appSearcher = new ManagementObjectSearcher(deploymentScope, appQuery);
                ManagementObjectCollection applications = appSearcher.Get();

                ManagementObject application = null;
                foreach (ManagementObject app in applications)
                {
                    application = app;
                    break;
                }

                if (application == null)
                {
                    Console.WriteLine($"WARNING: Application '{localizedDisplayName}' not found.");
                    continue;
                }

                application.Get();
                // Get SDMPackageXML property
                string sdmPackageXml = application ["SDMPackageXML"]?.ToString();

                // Serialize to XML
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(sdmPackageXml);

                // Sanitize file name
                string safeName = Regex.Replace(localizedDisplayName, @"[\\/:*?""<>|\r\n]+", "_");
                safeName = Regex.Replace(safeName, @"\s+", "_");
                string xmlFileName = $"{deploymentName}_{safeName}.xml";

                // Save XML to file
                xmlDoc.Save(xmlFileName);
                Console.WriteLine($"Information about application was written to file {xmlFileName}");

                // Extract information from XML
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("ab", "http://schemas.microsoft.com/SystemsCenterConfigurationManager/2011/03/ApplicationManifest");

                XmlNode appMgmtDigest = xmlDoc.SelectSingleNode("//ab:AppMgmtDigest", nsmgr);
                if (appMgmtDigest != null)
                {
                    string appName = appMgmtDigest.SelectSingleNode("ab:Name", nsmgr)?.InnerText;
                    string appVersion = appMgmtDigest.SelectSingleNode("ab:Version", nsmgr)?.InnerText;
                    string appPublisher = appMgmtDigest.SelectSingleNode("ab:Publisher", nsmgr)?.InnerText;
                    string appInstallDate = appMgmtDigest.SelectSingleNode("ab:InstallDate", nsmgr)?.InnerText;
                    string appInstallSource = appMgmtDigest.SelectSingleNode("ab:InstallSource", nsmgr)?.InnerText;

                    Console.WriteLine($"WARNING:        And that application name:                    '{appName}'");
                    Console.WriteLine($"WARNING:        And that application version:                 '{appVersion}'");
                    Console.WriteLine($"WARNING:        And that application publisher:               '{appPublisher}'");
                    Console.WriteLine($"WARNING:        And that application install date:            '{appInstallDate}'");
                    Console.WriteLine($"WARNING:        And that application install source:          '{appInstallSource}'");

                    // Get Install Arguments
                    try
                    {
                        XmlNodeList argNodes = xmlDoc.SelectNodes("//ab:DeploymentType/ab:Installer/ab:InstallAction/ab:Args/ab:Arg", nsmgr);
                        foreach (XmlNode argNode in argNodes)
                        {
                            string argName = argNode.Attributes ["Name"]?.Value;
                            string argText = argNode.InnerText;
                            Console.WriteLine($"WARNING:        And that application install arguments:       '{argName}' '{argText}'");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"WARNING:        Error getting install arguments: {ex.Message}");
                    }


                    // Get CommandLineArg
                    XmlNode commandLineArgNode = xmlDoc.SelectSingleNode("//ab:DeploymentType/ab:Installer/ab:InstallAction/ab:Args/ab:Arg[@Name='CommandLine']", nsmgr);
                    string commandLineArg = commandLineArgNode?.InnerText;

                    Console.WriteLine($"WARNING:        And that application command line:             '{commandLineArg}'");

                    // Get InstallCommandLine
                    string installCommandLine = xmlDoc.SelectSingleNode("//ab:DeploymentType/ab:Installer/ab:CustomData/ab:InstallCommandLine", nsmgr)?.InnerText;
                    Console.WriteLine($"WARNING:        And that application install command line:     '{installCommandLine}'");

                    // Get UninstallCommandLine
                    string uninstallCommandLine = xmlDoc.SelectSingleNode("//ab:DeploymentType/ab:Installer/ab:CustomData/ab:UninstallCommandLine", nsmgr)?.InnerText;
                    Console.WriteLine($"WARNING:        And that application uninstall command line:   '{uninstallCommandLine}'");
                }
                else
                {
                    Console.WriteLine("WARNING:        AppMgmtDigest node not found in XML.");
                }

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("=============================");
                Console.ResetColor();
            }
        }
    }
}
