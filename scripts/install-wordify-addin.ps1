[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ManifestPath = "..\manifest.xml",

    [Parameter(Mandatory = $false)]
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"

$resolvedManifestPath = $null
if (-not $Uninstall) {
    try {
        $resolvedManifestPath = (Resolve-Path -Path $ManifestPath).Path
    }
    catch {
        throw "Could not find manifest file at '$ManifestPath'. Provide a valid -ManifestPath."
    }
}

$wefFolder = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
if (-not (Test-Path -Path $wefFolder)) {
    New-Item -ItemType Directory -Path $wefFolder -Force | Out-Null
}

if ($Uninstall) {
    $manifestFileName = [System.IO.Path]::GetFileName($ManifestPath)
    $targetPath = Join-Path $wefFolder $manifestFileName

    if (Test-Path -Path $targetPath) {
        Remove-Item -Path $targetPath -Force
        Write-Host "Removed add-in manifest: $targetPath"
    }
    else {
        Write-Host "No installed manifest found at: $targetPath"
    }

    Write-Host "Done. Restart Word if it is currently running."
    exit 0
}

$manifestFileName = [System.IO.Path]::GetFileName($resolvedManifestPath)
$targetManifestPath = Join-Path $wefFolder $manifestFileName

Copy-Item -Path $resolvedManifestPath -Destination $targetManifestPath -Force

Write-Host "Installed Word add-in manifest to: $targetManifestPath"
Write-Host "Restart Word, then open Insert > My Add-ins to load the add-in."
