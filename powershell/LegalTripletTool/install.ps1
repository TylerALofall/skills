[CmdletBinding()]
param(
    [string]$SourcePath = (Join-Path $PSScriptRoot '.'),
    [switch]$CurrentUser,
    [switch]$Validate = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$moduleName = 'LegalTripletTool'

function Get-InstallPath {
    param([switch]$CurrentUser)

    if ($CurrentUser) {
        if ($IsWindows) {
            return Join-Path $HOME "Documents/PowerShell/Modules/$moduleName"
        }
        return Join-Path $HOME ".local/share/powershell/Modules/$moduleName"
    }

    $target = ($env:PSModulePath -split [IO.Path]::PathSeparator | Select-Object -First 1)
    if ([string]::IsNullOrWhiteSpace($target)) {
        throw 'Unable to determine PSModulePath target.'
    }
    return Join-Path $target $moduleName
}

$destination = Get-InstallPath -CurrentUser:$CurrentUser
New-Item -Path $destination -ItemType Directory -Force | Out-Null
Copy-Item -Path (Join-Path $SourcePath '*') -Destination $destination -Recurse -Force

Import-Module $moduleName -Force

if ($Validate) {
    $required = @(
        'Build-EvidenceSourceXml',
        'Get-LegalObjectCollection',
        'ConvertTo-LegalTripletSet',
        'Resolve-LegalSeedLinks',
        'Search-LegalTriplets',
        'Start-ComputerControlConsole'
    )

    $missing = @($required | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($missing.Count -gt 0) {
        throw "Installed but missing commands: $($missing -join ', ')"
    }
}

[pscustomobject]@{
    Module = $moduleName
    InstalledTo = $destination
    Version = (Get-Module $moduleName).Version.ToString()
    Validated = [bool]$Validate
}
