# .NET 8.0 Upgrade Plan

## Execution Steps

Execute steps below sequentially one by one in the order they are listed.

1. Validate that a .NET 8.0 SDK required for this upgrade is installed on the machine and if not, help to get it installed.
2. Ensure that the SDK version specified in global.json files is compatible with the .NET 8.0 upgrade.
3. Upgrade ListAllApps\ListAllApps.csproj to .NET 8.0
4. Upgrade DeploymentInfoLogger\DeploymentInfoLogger.csproj to .NET 8.0
5. Upgrade SCCMInfo\SCCMInfo.csproj to .NET 8.0

## Settings

This section contains settings and data used by execution steps.

### Excluded projects

| Project name | Description |
|:-------------|:-----------:|
| (none)       |             |

### Aggregate NuGet packages modifications across all projects

| Package Name              | Current Version | New Version | Description                                          |
|:--------------------------|:---------------:|:-----------:|:-----------------------------------------------------|
| System.Buffers            | 4.6.1           |             | Included in .NET 8.0 framework, remove reference     |
| System.Memory             | 4.6.3           |             | Included in .NET 8.0 framework, remove reference     |
| System.Numerics.Vectors   | 4.6.1           |             | Included in .NET 8.0 framework, remove reference     |

### Project upgrade details

#### ListAllApps\ListAllApps.csproj modifications

Project properties changes:
  - Target framework should be changed from `net48` to `net8.0`
  - Project file must be converted to SDK-style format

#### DeploymentInfoLogger\DeploymentInfoLogger.csproj modifications

Project properties changes:
  - Target framework should be changed from `net48` to `net8.0`
  - Project file must be converted to SDK-style format

#### SCCMInfo\SCCMInfo.csproj modifications

Project properties changes:
  - Target framework should be changed from `net48` to `net8.0`
  - Project file must be converted to SDK-style format

NuGet packages changes:
  - System.Buffers should be removed (included in .NET 8.0 platform)
  - System.Memory should be removed (included in .NET 8.0 platform)
  - System.Numerics.Vectors should be removed (included in .NET 8.0 platform)
