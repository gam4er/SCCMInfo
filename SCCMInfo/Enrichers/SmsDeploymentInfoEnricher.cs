using Spectre.Console;

using System;
using System.IO;
using System.Management;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace SCCMInfo.Enrichers
{
    internal sealed class SmsDeploymentInfoEnricher : IInstanceEnricher
    {
        public string ClassName => "SMS_DeploymentInfo";

        public void Enrich(
            ManagementBaseObject targetInstance,
            Table table,
            StringBuilder logMessage,
            ManagementScope scope)
        {
            string collectionID = targetInstance["CollectionID"]?.ToString();
            string targetName = targetInstance["TargetName"]?.ToString();
            string deploymentName = targetInstance["DeploymentName"]?.ToString();
            string collectionName = targetInstance["CollectionName"]?.ToString();

            table.AddRow("CollectionID", collectionID);
            table.AddRow("CollectionName", collectionName);

            string queryString = $"SELECT * FROM SMS_FullCollectionMembership WHERE CollectionID='{collectionID}'";
            ObjectQuery query = new ObjectQuery(queryString);
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

            ManagementObjectCollection collectionMembers = searcher.Get();

            var csvData = new StringBuilder();
            csvData.AppendLine("Name,ResourceID,ResourceType,ResourceDomainORWorkgroup,LastChangeTime,IsDirect");

            foreach (ManagementObject member in collectionMembers)
            {
                member.Get();
                string name = string.Empty;
                string resourceID = string.Empty;
                string resourceType = string.Empty;
                string domain = string.Empty;
                string isClient = string.Empty;
                string isActive = string.Empty;
                string isApproved = string.Empty;
                string isBlocked = string.Empty;

                try { name = member["Name"]?.ToString() ?? string.Empty; }
                catch { }

                try { resourceID = member["ResourceID"]?.ToString() ?? string.Empty; }
                catch { }

                try { resourceType = member["ResourceType"]?.ToString() ?? string.Empty; }
                catch { }

                try { domain = member["Domain"]?.ToString() ?? string.Empty; }
                catch { }

                try { isClient = member["IsClient"]?.ToString() ?? string.Empty; }
                catch { }

                try { isActive = member["IsActive"]?.ToString() ?? string.Empty; }
                catch { }

                try { isApproved = member["IsApproved"]?.ToString() ?? string.Empty; }
                catch { }

                try { isBlocked = member["IsBlocked"]?.ToString() ?? string.Empty; }
                catch { }

                csvData.AppendLine($"{name},{resourceID},{resourceType},{domain},{isClient},{isActive},{isApproved},{isBlocked}");
            }

            string csvFileName = $"{deploymentName}_{collectionID}.csv";
            File.WriteAllText(csvFileName, csvData.ToString());

            int membersCount = collectionMembers.Count;
            table.AddRow("CollectionMembersCount", membersCount.ToString());
            table.AddRow("CollectionMembersFile", csvFileName);

            logMessage.AppendLine($"Collection members info saved to file {csvFileName}, total members: {membersCount}");
            logMessage.AppendLine($"Collection Name: {collectionName}, ID: {collectionID}");

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
                    string sdmPackageXml = application["SDMPackageXML"]?.ToString();
                    logMessage.AppendLine($"sdmPackageXml gotted");
                    try
                    {
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.LoadXml(sdmPackageXml);

                        string safeName = Regex.Replace(targetName ?? string.Empty, @"[\\/:*?""<>|\r\n]+", "_");
                        safeName = Regex.Replace(safeName, @"\s+", "_");
                        string xmlFileName = $"{deploymentName}_{safeName}.xml";

                        xmlDoc.Save(xmlFileName);
                        logMessage.AppendLine($"Application info saved to file {xmlFileName}");

                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                        nsmgr.AddNamespace("ab", "http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest");

                        XmlNode appMgmtDigest = xmlDoc.SelectSingleNode("//ab:AppMgmtDigest", nsmgr);
                        if (appMgmtDigest != null)
                        {
                            string appName = appMgmtDigest.SelectSingleNode("ab:Name", nsmgr)?.InnerText ?? string.Empty;
                            string appVersion = appMgmtDigest.SelectSingleNode("ab:Version", nsmgr)?.InnerText ?? string.Empty;
                            string appPublisher = appMgmtDigest.SelectSingleNode("ab:Publisher", nsmgr)?.InnerText ?? string.Empty;
                            string appInstallDate = appMgmtDigest.SelectSingleNode("ab:InstallDate", nsmgr)?.InnerText ?? string.Empty;
                            string appInstallSource = appMgmtDigest.SelectSingleNode("ab:InstallSource", nsmgr)?.InnerText ?? string.Empty;

                            table.AddRow("ApplicationName", appName);
                            table.AddRow("ApplicationVersion", appVersion);
                            table.AddRow("ApplicationPublisher", appPublisher);
                            table.AddRow("ApplicationInstallDate", appInstallDate);
                            table.AddRow("ApplicationInstallSource", appInstallSource);
                            table.AddRow("ApplicationXMLFile", xmlFileName);
                            logMessage.AppendLine("Application info saved to vars");

                            try
                            {
                                XmlNodeList argNodes = xmlDoc.SelectNodes("//ab:DeploymentType/ab:Installer/ab:InstallAction/ab:Args/ab:Arg", nsmgr);
                                foreach (XmlNode argNode in argNodes)
                                {
                                    string argName = argNode.Attributes?["Name"]?.Value;
                                    string argText = argNode.InnerText;
                                    logMessage.AppendLine($"And that application install arguments:       '{argName}' '{argText}'");
                                    table.AddRow(argName, argText);
                                }
                                logMessage.AppendLine("Application Install Arguments saved");
                            }
                            catch
                            {
                                logMessage.AppendLine("Error while processing properties SMS_ApplicationLatest: InstallArguments not found");
                            }

                            try
                            {
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
    }
}
