# SCCMInfo

*A Windows service and console utility that monitors critical Configuration Manager (SCCM) events through WMI and writes rich diagnostics for administrators.*

[Читать на русском](README.ru.md)

## Table of contents
- [Overview](#overview)
- [Features](#features)
- [Repository layout](#repository-layout)
- [Prerequisites](#prerequisites)
- [Building](#building)
  - [Visual Studio](#visual-studio)
  - [Command line](#command-line)
- [Usage](#usage)
  - [Interactive console mode](#interactive-console-mode)
  - [Running as a Windows service](#running-as-a-windows-service)
- [Logging](#logging)
- [Troubleshooting](#troubleshooting)
- [License](#license)

## Overview
SCCMInfo connects to the `ROOT\\SMS` namespace of the current SCCM site, subscribes
to the creation of key objects (such as `SMS_DeploymentInfo` and
`SMS_CombinedDeviceResources`), enriches the captured payloads, and records them to
both a structured console table and persistent logs. The tool can run
interactively for diagnostics or be installed as a Windows service for continuous
monitoring.

## Features
- Automatic detection of the local management point and SCCM site code.
- Subscription to SCCM WMI events with pluggable enrichers that append contextual
  data before logging.
- Console output rendered with Spectre.Console for quick situational awareness
  during interactive sessions.
- Persistent logging to `ProcessInfoLog.txt` alongside the executable and, when
  permitted, to the Windows Application event log.
- Optional Windows service mode for unattended monitoring.

## Repository layout
| Path | Description |
| --- | --- |
| `SCCMInfo/` | Source code for the console application and Windows service. |
| `SCCMInfo.sln` | Visual Studio solution targeting .NET Framework 4.8. |
| `LICENSE.txt` | Project license. |

## Prerequisites
- Windows 10/11 or Windows Server with access to Configuration Manager
  infrastructure.
- .NET Framework 4.8 with Visual Studio 2019+ or the corresponding Build Tools.
- Local administrator rights and permissions to access WMI namespaces
  `ROOT\\SMS` and `ROOT\\ccm`.
- Access to the SMS Provider (locally or remotely) and permission to write to the
  Windows Application event log.

## Building
Restore NuGet packages before compiling (e.g., `nuget restore SCCMInfo.sln`).

### Visual Studio
1. Open `SCCMInfo.sln` in Visual Studio 2019 or later.
2. Choose the desired configuration (`Debug` or `Release`).
3. Build the solution via **Build → Build Solution**.

### Command line
Use the **Developer Command Prompt for VS** or an environment with MSBuild:

```powershell
nuget restore SCCMInfo.sln
msbuild SCCMInfo.sln /t:Build /p:Configuration=Release
```

The compiled binaries are produced in `SCCMInfo/bin/<Configuration>`.

## Usage
### Interactive console mode
1. Ensure the SCCM client is installed and that the machine can reach the target
   `ROOT\\ccm` and `ROOT\\SMS` namespaces.
2. Run `SCCMInfo.exe` without parameters to start interactive monitoring. WMI
   events are displayed in the console and written to `ProcessInfoLog.txt`
   alongside the executable.
3. Review console output to verify enrichers are augmenting the captured
   instances as expected.

### Running as a Windows service
1. Open an elevated PowerShell or Command Prompt window.
2. Install the service:
   ```powershell
   SCCMInfo.exe --install
   ```
   The installation registers a service named `SCCMInfo` and starts it
   automatically.
3. Manage the service as needed:
   - Start: `sc start SCCMInfo`
   - Stop: `sc stop SCCMInfo`
   - Remove after stopping: `sc delete SCCMInfo`
4. To validate service behavior without installing, execute:
   ```powershell
   SCCMInfo.exe --service
   ```
   This launches the service handlers in-process for debugging scenarios.

## Logging
SCCMInfo records activity to multiple sinks:
- **`ProcessInfoLog.txt`** — stored next to `SCCMInfo.exe`, containing detailed
  event payloads and errors.
- **Windows Application log** — mirrors key messages with event IDs 2001 and 2002
  under the source name `SCCMInfo` (created automatically if missing).
- **Console output** — presents Spectre.Console tables and verbose status updates
  when running interactively.

## Troubleshooting
- Confirm the executing account has permission to query the SCCM WMI namespaces
  and to create event log sources.
- If remote SMS Provider access fails, run SCCMInfo under a domain account with
  appropriate rights or establish a trusted connection.
- Check `ProcessInfoLog.txt` for unhandled exceptions and enrichers that fail to
  attach additional context.
- When running as a service, ensure the log file directory is writable by the
  service account (Local System by default).

## License
Distributed under the terms of [LICENSE.txt](LICENSE.txt).
