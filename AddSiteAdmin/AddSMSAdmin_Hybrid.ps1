[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$UserIdentity,

    [string]$SiteServer = "sccm.gam.click",
    [string]$SiteCode   = "GAM",

    # hardcoded known-good IDs from your environment
    [string]$RoleID = "SMS0001R",          # Full Administrator
    [string]$SecurityScopeID = "SMS00ALL", # All
    [string[]]$CollectionIDs = @("SMS00001", "SMS00004"), # All Systems, All Users and User Groups

    [string]$LogPath = (Join-Path (Get-Location) ("AddSMSAdmin_Hybrid_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Management
Add-Type -AssemblyName System.DirectoryServices

$Utf8Bom = [System.Text.UTF8Encoding]::new($true)

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR")][string]$Level = "INFO"
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line
    [System.IO.File]::AppendAllText($script:LogPath, $line + [Environment]::NewLine, $script:Utf8Bom)
}

function Escape-LdapFilterValue {
    param([Parameter(Mandatory = $true)][string]$Value)

    $sb = [System.Text.StringBuilder]::new()
    foreach ($ch in $Value.ToCharArray()) {
        switch ($ch) {
            '\' { [void]$sb.Append('\5c') }
            '*' { [void]$sb.Append('\2a') }
            '(' { [void]$sb.Append('\28') }
            ')' { [void]$sb.Append('\29') }
            ([char]0) { [void]$sb.Append('\00') }
            default { [void]$sb.Append($ch) }
        }
    }
    $sb.ToString()
}

function Escape-WqlString {
    param([Parameter(Mandatory = $true)][string]$Value)
    $Value.Replace("'", "''")
}

function Resolve-AdUser {
    param([Parameter(Mandatory = $true)][string]$Identity)

    if ($Identity -match '^CN=.+?,') {
        return [pscustomobject]@{
            DistinguishedName = $Identity
            NtAccount         = $null
        }
    }

    $rootDse    = [System.DirectoryServices.DirectoryEntry]::new("LDAP://RootDSE")
    $defaultNc  = [string]$rootDse.Properties["defaultNamingContext"][0]
    $searchRoot = [System.DirectoryServices.DirectoryEntry]::new("LDAP://$defaultNc")
    $searcher   = [System.DirectoryServices.DirectorySearcher]::new($searchRoot)

    $null = $searcher.PropertiesToLoad.Add("distinguishedName")
    $null = $searcher.PropertiesToLoad.Add("objectSid")

    if ($Identity -match '^[^\\]+\\[^\\]+$') {
        $sam = ($Identity -split '\\', 2)[1]
        $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$(Escape-LdapFilterValue $sam)))"
    }
    elseif ($Identity -match '@') {
        $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(userPrincipalName=$(Escape-LdapFilterValue $Identity)))"
    }
    else {
        $searcher.Filter = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=$(Escape-LdapFilterValue $Identity)))"
    }

    $result = $searcher.FindOne()
    if (-not $result) {
        throw "Пользователь '$Identity' не найден в AD."
    }

    $dn = [string]$result.Properties["distinguishedName"][0]

    $ntAccount = $null
    if ($result.Properties["objectSid"].Count -gt 0) {
        $sidBytes = [byte[]]$result.Properties["objectSid"][0]
        $sid = [System.Security.Principal.SecurityIdentifier]::new($sidBytes, 0)
        $ntAccount = $sid.Translate([System.Security.Principal.NTAccount]).Value
    }

    [pscustomobject]@{
        DistinguishedName = $dn
        NtAccount         = $ntAccount
    }
}

function New-WmiScope {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$NamespacePath
    )

    $options = [System.Management.ConnectionOptions]::new()
    $options.EnablePrivileges = $true
    $options.Impersonation = [System.Management.ImpersonationLevel]::Impersonate
    $options.Authentication = [System.Management.AuthenticationLevel]::PacketPrivacy

    $scope = [System.Management.ManagementScope]::new("\\$ComputerName\$NamespacePath", $options)
    $scope.Connect()
    return $scope
}

function New-EmbeddedPermission {
    param(
        [Parameter(Mandatory = $true)][System.Management.ManagementScope]$Scope,
        [Parameter(Mandatory = $true)][string]$RoleID,
        [Parameter(Mandatory = $true)][string]$CategoryID,
        [Parameter(Mandatory = $true)][uint32]$CategoryTypeID
    )

    $permClass = [System.Management.ManagementClass]::new(
        $Scope,
        [System.Management.ManagementPath]::new("SMS_APermission"),
        $null
    )

    $perm = $permClass.CreateInstance()
    $perm["RoleID"] = $RoleID
    $perm["CategoryID"] = $CategoryID
    $perm["CategoryTypeID"] = $CategoryTypeID
    return $perm
}

try {
    [System.IO.File]::WriteAllText($LogPath, "", $Utf8Bom)

    Write-Log "Старт."
    Write-Log "Host: $((Get-Process -Id $PID).Path)"
    Write-Log "PowerShell: $($PSVersionTable.PSVersion) / Edition=$($PSVersionTable.PSEdition)"
    Write-Log "Лог: $LogPath"

    $resolved = Resolve-AdUser -Identity $UserIdentity
    Write-Log "Resolved DN: $($resolved.DistinguishedName)"
    if ($resolved.NtAccount) {
        Write-Log "Resolved NTAccount: $($resolved.NtAccount)"
    }

    $opt = New-CimSessionOption -Protocol Dcom
    $cim = New-CimSession -ComputerName $SiteServer -SessionOption $opt
    $ns = "root\SMS\site_$SiteCode"

    # READ ONLY THROUGH WORKING CIM
    Write-Log "Looking up role by RoleID='$RoleID'..."
    $role = Get-CimInstance -CimSession $cim -Namespace $ns -Query "SELECT RoleID, RoleName FROM SMS_Role WHERE RoleID='$(Escape-WqlString $RoleID)'"

    if (-not $role) {
        Write-Log "RoleID '$RoleID' not found directly. Trying RoleName='Full Administrator'..." "WARN"
        $role = Get-CimInstance -CimSession $cim -Namespace $ns -Query "SELECT RoleID, RoleName FROM SMS_Role WHERE RoleName='Full Administrator'"
    }

    if (-not $role) {
        $allRoles = @(Get-CimInstance -CimSession $cim -Namespace $ns -Query "SELECT RoleID, RoleName FROM SMS_Role")
        Write-Log "Available roles ($($allRoles.Count) total):" "WARN"
        foreach ($r in $allRoles) {
            Write-Log "  RoleID=$($r.RoleID)  RoleName=$($r.RoleName)" "WARN"
        }
        throw "Role not found by RoleID='$RoleID' or RoleName='Full Administrator'. See available roles above."
    }

    $RoleID = $role.RoleID
    Write-Log "Role resolved: $($role.RoleID) / $($role.RoleName)"

    Write-Log "Looking up security scope by CategoryID='$SecurityScopeID'..."
    $scopeObj = Get-CimInstance -CimSession $cim -Namespace $ns -Query "SELECT CategoryID, CategoryName FROM SMS_SecuredCategory WHERE CategoryID='$(Escape-WqlString $SecurityScopeID)'"
    if (-not $scopeObj) {
        $allScopes = @(Get-CimInstance -CimSession $cim -Namespace $ns -Query "SELECT CategoryID, CategoryName FROM SMS_SecuredCategory")
        Write-Log "Available security scopes ($($allScopes.Count) total):" "WARN"
        foreach ($s in $allScopes) {
            Write-Log "  CategoryID=$($s.CategoryID)  CategoryName=$($s.CategoryName)" "WARN"
        }
        throw "SecurityScopeID '$SecurityScopeID' not found. See available scopes above."
    }
    Write-Log "Scope resolved: $($scopeObj.CategoryID) / $($scopeObj.CategoryName)"

    $collectionObjs = foreach ($cid in $CollectionIDs) {
        Write-Log "Looking up collection by CollectionID='$cid'..."
        $c = Get-CimInstance -CimSession $cim -Namespace $ns -Query "SELECT CollectionID, Name FROM SMS_Collection WHERE CollectionID='$(Escape-WqlString $cid)'"
        if (-not $c) {
            $allCols = @(Get-CimInstance -CimSession $cim -Namespace $ns -Query "SELECT CollectionID, Name FROM SMS_Collection")
            Write-Log "Available collections ($($allCols.Count) total):" "WARN"
            foreach ($col in $allCols) {
                Write-Log "  CollectionID=$($col.CollectionID)  Name=$($col.Name)" "WARN"
            }
            throw "CollectionID '$cid' not found. See available collections above."
        }
        Write-Log "Collection resolved: $($c.CollectionID) / $($c.Name)"
        $c
    }

    $escapedDn = Escape-WqlString $resolved.DistinguishedName
    $existing = Get-CimInstance -CimSession $cim -Namespace $ns -Query "SELECT AdminID, LogonName, DistinguishedName FROM SMS_Admin WHERE DistinguishedName='$escapedDn'"

    if ($existing) {
        Write-Log "SMS_Admin already exists. AdminID=$($existing.AdminID); LogonName=$($existing.LogonName)" "WARN"
        return
    }

    Write-Log "SMS_Admin not found. Creating via System.Management Put()."

    # WRITE VIA WMI/PUT (CIM does not support CreateInstance for SMS_Admin)
    $scope = New-WmiScope -ComputerName $SiteServer -NamespacePath $ns

    # Build permissions array: 1 security scope (CategoryTypeID=29) + N collections (CategoryTypeID=1)
    $permissions = New-Object 'System.Collections.Generic.List[System.Management.ManagementBaseObject]'

    # SecuredScope (29)
    Write-Log "Adding permission: RoleID=$RoleID  CategoryID=$SecurityScopeID  CategoryTypeID=29"
    $permissions.Add(
        (New-EmbeddedPermission -Scope $scope -RoleID $RoleID -CategoryID $SecurityScopeID -CategoryTypeID 29)
    ) | Out-Null

    # Collections (1)
    foreach ($cid in $CollectionIDs) {
        Write-Log "Adding permission: RoleID=$RoleID  CategoryID=$cid  CategoryTypeID=1"
        $permissions.Add(
            (New-EmbeddedPermission -Scope $scope -RoleID $RoleID -CategoryID $cid -CategoryTypeID 1)
        ) | Out-Null
    }

    $adminClass = [System.Management.ManagementClass]::new(
        $scope,
        [System.Management.ManagementPath]::new("SMS_Admin"),
        $null
    )

    $newAdmin = $adminClass.CreateInstance()

    # Minimal required payload
    $newAdmin["DistinguishedName"] = $resolved.DistinguishedName
    if ($resolved.NtAccount) {
        $newAdmin["LogonName"] = $resolved.NtAccount
    }
    $newAdmin["Permissions"] = [System.Management.ManagementBaseObject[]]$permissions.ToArray()

    Write-Log "SMS_Admin instance prepared: DN=$($resolved.DistinguishedName)  LogonName=$($resolved.NtAccount)  Permissions=$($permissions.Count)"

    if ($PSCmdlet.ShouldProcess($resolved.DistinguishedName, "Create SMS_Admin")) {
        Write-Log "Calling Put()..."
        try {
            $path = $newAdmin.Put()
            Write-Log "Put() completed. Path=$path"
        }
        catch {
            Write-Log "Put() failed: $($_.Exception.Message)" "ERROR"
            Write-Log "Full exception: $($_.Exception.ToString())" "ERROR"
            throw
        }
    }

    Start-Sleep -Seconds 2

    $verify = Get-CimInstance -CimSession $cim -Namespace $ns -Query "SELECT AdminID, LogonName, DistinguishedName FROM SMS_Admin WHERE DistinguishedName='$escapedDn'"
    if (-not $verify) {
        throw "After Put(), SMS_Admin still not found."
    }

    Write-Log "Verification OK."
    Write-Log "AdminID: $($verify.AdminID)"
    Write-Log "LogonName: $($verify.LogonName)"
    Write-Log "DistinguishedName: $($verify.DistinguishedName)"

    # Read back permissions via Get-WmiObject (lazy embedded array SMS_Admin.Permissions
    # is not returned by CIM SELECT in SCCM 2012 R2 — WMI fetches the full instance)
    try {
        $wmiAdmin = Get-WmiObject -ComputerName $SiteServer -Namespace $ns `
            -Class SMS_Admin -Filter "AdminID=$($verify.AdminID)"
        if ($wmiAdmin -and $wmiAdmin.Permissions) {
            Write-Log "Permissions assigned ($($wmiAdmin.Permissions.Count) total):"
            foreach ($p in $wmiAdmin.Permissions) {
                Write-Log ("  RoleID={0}; CategoryID={1}; CategoryTypeID={2}" -f $p.RoleID, $p.CategoryID, $p.CategoryTypeID)
            }
        }
        else {
            Write-Log "Permissions property is empty or unavailable via WMI on this SCCM version." "WARN"
        }
    }
    catch {
        $ex = $_.Exception
        $hresultVal = [int64]4294967295 -band [int64]$ex.HResult
        Write-Log "Could not read back Permissions from SMS_Admin: $($ex.Message)  HRESULT=0x$($hresultVal.ToString('X8'))" "WARN"
    }

    Write-Log "Done."
}
catch {
    Write-Log $_.Exception.ToString() "ERROR"
    throw
}
finally {
    if ($cim) {
        $cim | Remove-CimSession -ErrorAction SilentlyContinue
    }
}