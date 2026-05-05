# .NET 8.0 Upgrade Report

## Project target framework modifications

| Project name                                              | Old Target Framework | New Target Framework | Commits                              |
|:----------------------------------------------------------|:--------------------:|:--------------------:|:-------------------------------------|
| ListAllApps\ListAllApps.csproj                            | net48                | net8.0               | c787591e, dd3a3def, ae0b4cd2         |
| DeploymentInfoLogger\DeploymentInfoLogger.csproj          | net48                | net8.0               | 5f47c9f4, 4271d95f, b2381991         |
| SCCMInfo\SCCMInfo.csproj                                  | net48                | net8.0               | 0e687b2c, 26b25aa7, 971341f5         |

## NuGet Packages

| Package Name                          | Old Version | New Version | Commit Id |
|:--------------------------------------|:-----------:|:-----------:|:----------|
| System.Buffers                        | 4.6.1       | removed     | 0e687b2c  |
| System.Configuration.ConfigurationManager |        | 10.0.7      | b2381991  |
| System.Diagnostics.EventLog           |             | 8.0.0       | (manual)  |
| System.Management (NuGet)             | (assembly)  | 8.0.0       | ae0b4cd2  |
| System.Memory                         | 4.6.3       | removed     | 0e687b2c  |
| System.Numerics.Vectors               | 4.6.1       | removed     | 0e687b2c  |
| System.ServiceProcess.ServiceController |           | 8.0.0       | (manual)  |

## All commits

| Commit ID  | Description                                                |
|:-----------|:-----------------------------------------------------------|
| 94aba659   | Commit upgrade plan                                        |
| c787591e   | Move assembly metadata from AssemblyInfo.cs to .csproj     |
| dd3a3def   | Modernize ListAllApps.csproj to SDK-style and .NET 8       |
| ae0b4cd2   | Remove unused references from ListAllApps.csproj           |
| 5f47c9f4   | Migrate DeploymentInfoLogger.csproj to SDK-style and .NET 8 |
| 4271d95f   | Move assembly metadata to .csproj and clean AssemblyInfo.cs |
| b2381991   | Update DeploymentInfoLogger.csproj package references      |
| 0e687b2c   | Migrate project to SDK-style and update to .NET 8          |
| 26b25aa7   | Clean up references in SCCMInfo.csproj                     |
| 971341f5   | Move assembly metadata from AssemblyInfo.cs to csproj      |

## Project feature upgrades

### DeploymentInfoLogger\DeploymentInfoLogger.csproj

- `System.Management.Instrumentation` namespace is not available in .NET 8.0 — WMI instrumentation attributes (`ManagementEntityAttribute`, `ManagementKeyAttribute`, `ManagementCreateAttribute`, `WmiConfigurationAttribute`) were commented out. If WMI provider functionality is required, a compatible alternative must be found.
- `System.EnterpriseServices.Internal` (GAC install/uninstall via `RegistrationServices`) is not available in .NET 8.0 — code was commented out with TODO notes.
- `DefaultManagementInstaller` base class not available in .NET 8.0 — inheritance removed with a comment.
- `Publish.GacInstall` / `Publish.GacRemove` not available in .NET 8.0 — calls were commented out.

## Next steps

- Review commented-out WMI instrumentation code in `DeploymentInfoLogger` and decide if WMI provider functionality is still needed — if so, explore alternative .NET 8.0-compatible libraries.
- Review commented-out GAC installation code in `DeploymentInfoLogger` and replace with a modern alternative if needed.
- Push branch `upgrade-to-NET8` to remote and create a Pull Request to merge changes into `dev`.
