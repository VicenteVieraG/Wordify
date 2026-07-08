[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
param(
    [Parameter(Mandatory = $false)]
    [string]$VbaPath,

    [Parameter(Mandatory = $false)]
    [string]$RepositoryUrl = "https://github.com/VicenteVieraG/Wordify.git",

    [Parameter(Mandatory = $false)]
    [string]$ClonePath = (Join-Path $HOME "Wordify"),

    [Parameter(Mandatory = $false)]
    [switch]$SkipSourceBootstrap,

    [Parameter(Mandatory = $false)]
    [switch]$SkipKeybindings,

    [Parameter(Mandatory = $false)]
    [switch]$Uninstall,

    [Parameter(Mandatory = $false)]
    [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$WdKeyCategoryMacro = 2
$WdDoNotSaveChanges = 0
$VbextStandardModule = 1

$ModuleSpecs = @(
    [pscustomobject]@{ Name = "modSpanishNumbers"; File = "modSpanishNumbers.bas" }
    [pscustomobject]@{ Name = "modConversions";    File = "modConversions.bas" }
    [pscustomobject]@{ Name = "modValidation";     File = "modValidation.bas" }
    [pscustomobject]@{ Name = "modTokenDetection"; File = "modTokenDetection.bas" }
    [pscustomobject]@{ Name = "modDeletion";       File = "modDeletion.bas" }
    [pscustomobject]@{ Name = "modUI";             File = "modUI.bas" }
    [pscustomobject]@{ Name = "modCommands";       File = "modCommands.bas" }
    [pscustomobject]@{ Name = "modSetup";          File = "modSetup.bas" }
)

$KeyBindings = @(
    [pscustomobject]@{ Command = "Wordify_F2_Time";                   Name = "F2";       Keys = @(113) }
    [pscustomobject]@{ Command = "Wordify_F3_Spell";                  Name = "F3";       Keys = @(114) }
    [pscustomobject]@{ Command = "Wordify_ShiftF3_SpellUpper";        Name = "Shift+F3"; Keys = @(256, 114) }
    [pscustomobject]@{ Command = "Wordify_F4_LandMeasure";            Name = "F4";       Keys = @(115) }
    [pscustomobject]@{ Command = "Wordify_F5_Number";                 Name = "F5";       Keys = @(116) }
    [pscustomobject]@{ Command = "Wordify_F6_CurrencyMXN";            Name = "F6";       Keys = @(117) }
    [pscustomobject]@{ Command = "Wordify_F7_LinearMeters";           Name = "F7";       Keys = @(118) }
    [pscustomobject]@{ Command = "Wordify_F8_SquareMeters";           Name = "F8";       Keys = @(119) }
    [pscustomobject]@{ Command = "Wordify_F9_DeleteToken";            Name = "F9";       Keys = @(120) }
    [pscustomobject]@{ Command = "Wordify_F10_DeleteToVisibleLineEnd"; Name = "F10";      Keys = @(121) }
    [pscustomobject]@{ Command = "Wordify_F11_DeleteVisibleLine";      Name = "F11";      Keys = @(122) }
)

function Add-ScoopToProcessPath {
    $candidatePaths = New-Object System.Collections.Generic.List[string]

    $scoopRoot = $env:SCOOP
    if ([string]::IsNullOrWhiteSpace($scoopRoot)) {
        $scoopRoot = Join-Path $HOME "scoop"
    }

    [void]$candidatePaths.Add((Join-Path $scoopRoot "shims"))

    $globalScoopRoot = $env:SCOOP_GLOBAL
    if ([string]::IsNullOrWhiteSpace($globalScoopRoot)) {
        $globalScoopRoot = Join-Path $env:ProgramData "scoop"
    }

    [void]$candidatePaths.Add((Join-Path $globalScoopRoot "shims"))

    $currentPaths = @($env:PATH -split [System.IO.Path]::PathSeparator)
    $pathsToPrepend = New-Object System.Collections.Generic.List[string]

    foreach ($candidatePath in $candidatePaths) {
        if ([string]::IsNullOrWhiteSpace($candidatePath)) {
            continue
        }

        if (-not (Test-Path -LiteralPath $candidatePath -PathType Container)) {
            continue
        }

        $alreadyPresent = $false
        foreach ($currentPath in $currentPaths) {
            if ($candidatePath.Equals($currentPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                $alreadyPresent = $true
                break
            }
        }

        if (-not $alreadyPresent) {
            [void]$pathsToPrepend.Add($candidatePath)
        }
    }

    if ($pathsToPrepend.Count -gt 0) {
        $env:PATH = (($pathsToPrepend + $currentPaths) -join [System.IO.Path]::PathSeparator)
    }
}

function Get-ExecutableCommandPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $command = Get-Command -Name $Name -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $command) {
        return $null
    }

    if (-not [string]::IsNullOrWhiteSpace($command.Source)) {
        return $command.Source
    }

    return $command.Name
}

function Test-ProcessIsElevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Invoke-CommandOutputForHost {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $output = @(& $ScriptBlock)
    foreach ($item in $output) {
        if ($item -is [System.Management.Automation.ErrorRecord]) {
            Write-Host $item.ToString()
        }
        else {
            Write-Host $item
        }
    }
}

function Set-ScoopInstallerExecutionPolicy {
    $permissivePolicies = @("Bypass", "RemoteSigned", "Unrestricted")
    $effectivePolicy = [string](Get-ExecutionPolicy)
    if ($permissivePolicies -contains $effectivePolicy) {
        Write-Host "PowerShell execution policy is $effectivePolicy for this process; continuing."
        return
    }

    try {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop
    }
    catch {
        $effectivePolicy = [string](Get-ExecutionPolicy)
        $isPolicyOverride = $_.FullyQualifiedErrorId -like "ExecutionPolicyOverride,*"
        if ($isPolicyOverride -and ($permissivePolicies -contains $effectivePolicy)) {
            Write-Host "PowerShell execution policy is $effectivePolicy because of a more specific scope; continuing."
            return
        }

        throw
    }
}

function Install-Scoop {
    Add-ScoopToProcessPath

    $scoopCommand = Get-ExecutableCommandPath -Name "scoop"
    if ($null -ne $scoopCommand) {
        Write-Host "Scoop is already installed: $scoopCommand"
        return $scoopCommand
    }

    if (-not $PSCmdlet.ShouldProcess("Scoop", "Install package manager for current user")) {
        return $null
    }

    Set-ScoopInstallerExecutionPolicy

    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $installerScript = (Invoke-WebRequest -UseBasicParsing -Uri "https://get.scoop.sh").Content

    $installerArguments = @()
    if (Test-ProcessIsElevated) {
        Write-Host "PowerShell is elevated; installing Scoop with -RunAsAdmin."
        $installerArguments += "-RunAsAdmin"
    }

    $Global:LASTEXITCODE = 0
    $installerBlock = [scriptblock]::Create($installerScript)
    Invoke-CommandOutputForHost -ScriptBlock { & $installerBlock @installerArguments }
    if ($LASTEXITCODE -ne 0) {
        throw "Scoop installer failed. Exit code: $LASTEXITCODE"
    }

    Add-ScoopToProcessPath

    $scoopCommand = Get-ExecutableCommandPath -Name "scoop"
    if ($null -eq $scoopCommand) {
        throw "Scoop was installed, but its command is not available on PATH. Open a new PowerShell session and rerun this script."
    }

    Write-Host "Installed Scoop: $scoopCommand"
    return $scoopCommand
}

function Install-Git {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ScoopCommand
    )

    Add-ScoopToProcessPath

    $gitCommand = Get-ExecutableCommandPath -Name "git"
    if ($null -ne $gitCommand) {
        Write-Host "Git is already installed: $gitCommand"
        return $gitCommand
    }

    if ([string]::IsNullOrWhiteSpace($ScoopCommand)) {
        throw "Git is required to clone Wordify, but Scoop is not available to install it."
    }

    if (-not $PSCmdlet.ShouldProcess("git", "Install Git with Scoop")) {
        return $null
    }

    Invoke-CommandOutputForHost -ScriptBlock { & $ScoopCommand install git }
    if ($LASTEXITCODE -ne 0) {
        throw "Scoop failed to install Git. Exit code: $LASTEXITCODE"
    }

    Add-ScoopToProcessPath

    $gitCommand = Get-ExecutableCommandPath -Name "git"
    if ($null -eq $gitCommand) {
        throw "Git was installed, but its command is not available on PATH. Open a new PowerShell session and rerun this script."
    }

    Write-Host "Installed Git: $gitCommand"
    return $gitCommand
}

function Resolve-AbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [System.IO.Path]::GetFullPath($Path)
}

function Test-WordifySourceRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    $vbaPath = Join-Path $Root "vba"
    $manifestPath = Join-Path $Root "manifest.xml"

    return (
        (Test-Path -LiteralPath $vbaPath -PathType Container) -and
        (Test-Path -LiteralPath $manifestPath -PathType Leaf)
    )
}

function Get-BundledWordifySourceRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptRoot
    )

    $candidateRoots = New-Object System.Collections.Generic.List[string]
    [void]$candidateRoots.Add($ScriptRoot)

    $parentRoot = Split-Path -Parent $ScriptRoot
    if (-not [string]::IsNullOrWhiteSpace($parentRoot)) {
        [void]$candidateRoots.Add($parentRoot)
    }

    $seen = @{}
    foreach ($candidateRoot in $candidateRoots) {
        if ([string]::IsNullOrWhiteSpace($candidateRoot)) {
            continue
        }

        $resolvedRoot = Resolve-AbsolutePath -Path $candidateRoot
        $key = $resolvedRoot.ToUpperInvariant()
        if ($seen.ContainsKey($key)) {
            continue
        }

        $seen[$key] = $true
        if (Test-WordifySourceRoot -Root $resolvedRoot) {
            return $resolvedRoot
        }
    }

    return $null
}

function Get-InstallerPackageRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptRoot
    )

    if ([string]::IsNullOrWhiteSpace($ScriptRoot)) {
        return $null
    }

    $packageRoot = $ScriptRoot
    $scriptRootName = Split-Path -Leaf $ScriptRoot
    if ($scriptRootName.Equals("scripts", [System.StringComparison]::OrdinalIgnoreCase)) {
        $parentRoot = Split-Path -Parent $ScriptRoot
        if (-not [string]::IsNullOrWhiteSpace($parentRoot)) {
            $packageRoot = $parentRoot
        }
    }

    return Resolve-AbsolutePath -Path $packageRoot
}

function Resolve-BootstrapClonePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PreferredPath,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$InstallerRoot,

        [Parameter(Mandatory = $true)]
        [bool]$ClonePathWasExplicit
    )

    $resolvedPreferredPath = Resolve-AbsolutePath -Path $PreferredPath
    if ($ClonePathWasExplicit) {
        return $resolvedPreferredPath
    }

    if (Test-WordifySourceRoot -Root $resolvedPreferredPath) {
        return $resolvedPreferredPath
    }

    if (-not (Test-Path -LiteralPath $resolvedPreferredPath -PathType Container)) {
        return $resolvedPreferredPath
    }

    $existingItems = @(Get-ChildItem -LiteralPath $resolvedPreferredPath -Force)
    if ($existingItems.Count -eq 0) {
        return $resolvedPreferredPath
    }

    $fallbackPath = Join-Path $resolvedPreferredPath ".wordify-source"
    if (-not [string]::IsNullOrWhiteSpace($InstallerRoot)) {
        $resolvedInstallerRoot = Resolve-AbsolutePath -Path $InstallerRoot
        if ($resolvedPreferredPath.Equals($resolvedInstallerRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Host "Installer package occupies the default clone folder. Using source clone path: $fallbackPath"
            return $fallbackPath
        }
    }

    Write-Host "Default clone folder exists but is not a Wordify source folder. Using source clone path: $fallbackPath"
    return $fallbackPath
}

function Initialize-WordifySourceCode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryUrl,

        [Parameter(Mandatory = $true)]
        [string]$ClonePath,

        [Parameter(Mandatory = $true)]
        [string]$GitCommand
    )

    if ([string]::IsNullOrWhiteSpace($RepositoryUrl)) {
        throw "RepositoryUrl cannot be empty."
    }

    if ([string]::IsNullOrWhiteSpace($ClonePath)) {
        throw "ClonePath cannot be empty."
    }

    $resolvedClonePath = Resolve-AbsolutePath -Path $ClonePath
    if (Test-WordifySourceRoot -Root $resolvedClonePath) {
        Write-Host "Using existing Wordify source checkout: $resolvedClonePath"
        return $resolvedClonePath
    }

    if (Test-Path -LiteralPath $resolvedClonePath -PathType Container) {
        $existingItems = @(Get-ChildItem -LiteralPath $resolvedClonePath -Force)
        if ($existingItems.Count -gt 0) {
            throw "Clone target exists but is not an empty folder or a Wordify checkout: $resolvedClonePath"
        }
    }
    else {
        $parentPath = Split-Path -Parent $resolvedClonePath
        if (-not [string]::IsNullOrWhiteSpace($parentPath)) {
            if ($PSCmdlet.ShouldProcess($parentPath, "Create source parent directory")) {
                [void](New-Item -ItemType Directory -Path $parentPath -Force)
            }
        }
    }

    if (-not $PSCmdlet.ShouldProcess($resolvedClonePath, "Clone $RepositoryUrl")) {
        return $resolvedClonePath
    }

    Invoke-CommandOutputForHost -ScriptBlock { & $GitCommand clone $RepositoryUrl $resolvedClonePath }
    if ($LASTEXITCODE -ne 0) {
        throw "Git failed to clone Wordify. Exit code: $LASTEXITCODE"
    }

    if (-not (Test-WordifySourceRoot -Root $resolvedClonePath)) {
        throw "The cloned repository does not look like a Wordify source checkout: $resolvedClonePath"
    }

    Write-Host "Cloned Wordify source checkout: $resolvedClonePath"
    return $resolvedClonePath
}

function Initialize-WordifyInstallSource {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryUrl,

        [Parameter(Mandatory = $true)]
        [string]$ClonePath
    )

    $resolvedClonePath = Resolve-AbsolutePath -Path $ClonePath
    if (Test-WordifySourceRoot -Root $resolvedClonePath) {
        Write-Host "Using existing Wordify source: $resolvedClonePath"
        return $resolvedClonePath
    }

    if (Test-Path -LiteralPath $resolvedClonePath -PathType Container) {
        $existingItems = @(Get-ChildItem -LiteralPath $resolvedClonePath -Force)
        if ($existingItems.Count -gt 0) {
            throw "Clone target exists but is not an empty folder or a Wordify source folder: $resolvedClonePath"
        }
    }

    Write-Host "Preparing Scoop and Git for Wordify source bootstrap."
    $scoopCommand = Install-Scoop
    $gitCommand = Install-Git -ScoopCommand $scoopCommand
    if ([string]::IsNullOrWhiteSpace($gitCommand)) {
        throw "Git is required to clone Wordify."
    }

    return Initialize-WordifySourceCode -RepositoryUrl $RepositoryUrl -ClonePath $ClonePath -GitCommand $gitCommand
}

function Get-NormalizedPathForComparison {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $trimChars = [char[]]("\\/")
    return (Resolve-AbsolutePath -Path $Path).TrimEnd($trimChars)
}

function Get-WordifySourceRemovalCandidates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClonePath,

        [Parameter(Mandatory = $true)]
        [bool]$ClonePathWasExplicit
    )

    $candidatePaths = New-Object System.Collections.Generic.List[string]
    $resolvedClonePath = Resolve-AbsolutePath -Path $ClonePath
    [void]$candidatePaths.Add($resolvedClonePath)

    if (-not $ClonePathWasExplicit) {
        [void]$candidatePaths.Add((Join-Path $resolvedClonePath ".wordify-source"))
    }

    $seen = @{}
    foreach ($candidatePath in $candidatePaths) {
        $resolvedCandidatePath = Resolve-AbsolutePath -Path $candidatePath
        $key = $resolvedCandidatePath.ToUpperInvariant()
        if ($seen.ContainsKey($key)) {
            continue
        }

        $seen[$key] = $true
        $resolvedCandidatePath
    }
}

function Assert-SafeWordifySourceRemovalTarget {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )

    $resolvedTargetPath = (Resolve-Path -LiteralPath $TargetPath).ProviderPath
    if (-not (Test-WordifySourceRoot -Root $resolvedTargetPath)) {
        throw "Refusing to remove source folder because it is not a recognized Wordify source folder: $resolvedTargetPath"
    }

    $normalizedTarget = Get-NormalizedPathForComparison -Path $resolvedTargetPath
    $normalizedRoot = (Get-NormalizedPathForComparison -Path ([System.IO.Path]::GetPathRoot($resolvedTargetPath)))
    if ($normalizedTarget.Equals($normalizedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove filesystem root as a source folder: $resolvedTargetPath"
    }

    $normalizedHome = Get-NormalizedPathForComparison -Path $HOME
    if ($normalizedTarget.Equals($normalizedHome, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove the user profile directory as a source folder: $resolvedTargetPath"
    }

    return $resolvedTargetPath
}

function Remove-WordifySourceFolders {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClonePath,

        [Parameter(Mandatory = $true)]
        [bool]$ClonePathWasExplicit
    )

    $removed = New-Object System.Collections.Generic.List[string]
    $candidatePaths = @(Get-WordifySourceRemovalCandidates -ClonePath $ClonePath -ClonePathWasExplicit $ClonePathWasExplicit)

    foreach ($candidatePath in $candidatePaths) {
        if (-not (Test-Path -LiteralPath $candidatePath -PathType Container)) {
            continue
        }

        if (-not (Test-WordifySourceRoot -Root $candidatePath)) {
            Write-Host "Skipping source cleanup because this folder is not a Wordify source folder: $candidatePath"
            continue
        }

        $safeTargetPath = Assert-SafeWordifySourceRemovalTarget -TargetPath $candidatePath
        if ($PSCmdlet.ShouldProcess($safeTargetPath, "Remove Wordify source folder")) {
            Remove-Item -LiteralPath $safeTargetPath -Recurse -Force
            [void]$removed.Add($safeTargetPath)
        }
    }

    return $removed
}

function Resolve-WordifyModuleFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [Parameter(Mandatory = $true)]
        [object[]]$Specs
    )

    $resolvedRoot = $null
    try {
        $resolvedRoot = (Resolve-Path -LiteralPath $Root).ProviderPath
    }
    catch {
        throw "Could not find the VBA module folder at '$Root'. Use -VbaPath to point at the Wordify vba folder."
    }

    foreach ($spec in $Specs) {
        $path = Join-Path $resolvedRoot $spec.File
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Missing required VBA module: $path"
        }

        $moduleName = Get-BasModuleName -Path $path
        if ($moduleName -ne $spec.Name) {
            throw "Module '$path' declares '$moduleName', but the installer expected '$($spec.Name)'."
        }

        [pscustomobject]@{
            Name = $spec.Name
            File = $spec.File
            Path = (Resolve-Path -LiteralPath $path).ProviderPath
        }
    }
}

function Get-BasModuleName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $match = Select-String -LiteralPath $Path -Pattern '^Attribute VB_Name = "([^"]+)"' -List
    if ($null -eq $match) {
        throw "Could not read the VBA module name from '$Path'."
    }

    return $match.Matches[0].Groups[1].Value
}

function Assert-WordAutomationIsAvailable {
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        throw "This installer requires Windows and Microsoft Word desktop."
    }
}

function New-WordApplication {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$ShowWindow
    )

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $ShowWindow
        $word.DisplayAlerts = 0
        return $word
    }
    catch {
        throw "Could not start Microsoft Word through COM automation. Confirm Word desktop is installed. $($_.Exception.Message)"
    }
}

function Get-NormalTemplateLockPath {
    $templatesPath = Join-Path $env:APPDATA "Microsoft\Templates"
    return (Join-Path $templatesPath '~$Normal.dotm')
}

function Assert-NormalTemplateIsAvailable {
    $wordProcesses = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
    if ($wordProcesses.Count -gt 0) {
        $processIds = $wordProcesses.Id -join ", "
        throw "Microsoft Word is currently running (WINWORD process id(s): $processIds). Close Word and rerun this installer so Normal.dotm can be updated."
    }

    $lockPath = Get-NormalTemplateLockPath
    if (Test-Path -LiteralPath $lockPath -PathType Leaf) {
        Remove-Item -LiteralPath $lockPath -Force
        Write-Host "Removed stale Normal.dotm lock file: $lockPath"
    }
}

function Assert-VbaProjectAccess {
    param(
        [Parameter(Mandatory = $true)]
        $Template
    )

    try {
        if ($null -eq $Template.VBProject) {
            throw "The template VBA project is not loaded."
        }

        [void]$Template.VBProject.VBComponents.Count
    }
    catch {
        $detail = $_.Exception.Message
        throw @"
Word blocked access to the VBA project. In Word, enable:
File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model

Then close Word and run this script again.

COM error: $detail
"@
    }
}

function Initialize-WordVbaProjectModel {
    param(
        [Parameter(Mandatory = $true)]
        $Word
    )

    $document = $null
    try {
        $document = $Word.Documents.Add()
        if ($null -eq $Word.VBE) {
            throw "Word did not expose the Visual Basic Editor automation object."
        }

        return $document
    }
    catch {
        if ($null -ne $document) {
            Close-WordDocumentNoSave -Document $document
        }

        throw "Could not initialize Word's VBA automation model. $($_.Exception.Message)"
    }
}

function Close-WordDocumentNoSave {
    param(
        [Parameter(Mandatory = $true)]
        $Document
    )

    $saveChanges = $script:WdDoNotSaveChanges
    [void]$Document.Close([ref]$saveChanges)
}

function Get-VbaComponent {
    param(
        [Parameter(Mandatory = $true)]
        $Components,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        return $Components.Item($Name)
    }
    catch {
        return $null
    }
}

function Remove-WordifyModules {
    param(
        [Parameter(Mandatory = $true)]
        $Template,

        [Parameter(Mandatory = $true)]
        [object[]]$Specs
    )

    $components = $Template.VBProject.VBComponents
    $removed = New-Object System.Collections.Generic.List[string]

    foreach ($spec in $Specs) {
        $component = Get-VbaComponent -Components $components -Name $spec.Name
        if ($null -eq $component) {
            continue
        }

        if ([int]$component.Type -ne $script:VbextStandardModule) {
            throw "Found a non-standard VBA component named '$($spec.Name)'. Refusing to remove it automatically."
        }

        if ($PSCmdlet.ShouldProcess($spec.Name, "Remove existing Wordify VBA module")) {
            [void]$components.Remove($component)
            [void]$removed.Add($spec.Name)
        }
    }

    return $removed
}

function Import-WordifyModules {
    param(
        [Parameter(Mandatory = $true)]
        $Template,

        [Parameter(Mandatory = $true)]
        [object[]]$Modules
    )

    $components = $Template.VBProject.VBComponents
    $imported = New-Object System.Collections.Generic.List[string]

    foreach ($module in $Modules) {
        if ($PSCmdlet.ShouldProcess($module.Path, "Import Wordify VBA module")) {
            [void]$components.Import($module.Path)
            [void]$imported.Add($module.Name)
        }
    }

    return $imported
}

function Clear-WordifyKeyBindings {
    param(
        [Parameter(Mandatory = $true)]
        $Word,

        [Parameter(Mandatory = $true)]
        $Template,

        [Parameter(Mandatory = $true)]
        [object[]]$Bindings
    )

    $lookup = @{}
    foreach ($binding in $Bindings) {
        $lookup[$binding.Command.ToUpperInvariant()] = $true
    }

    $Word.CustomizationContext = $Template
    $cleared = 0

    foreach ($binding in $Word.KeyBindings) {
        $command = $null
        try {
            $command = [string]$binding.Command
        }
        catch {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($command)) {
            continue
        }

        $macroName = $command
        if ($macroName.Contains(".")) {
            $macroName = $macroName.Split(".")[-1]
        }

        if ($lookup.ContainsKey($macroName.ToUpperInvariant())) {
            if ($PSCmdlet.ShouldProcess($command, "Clear Wordify keybinding")) {
                [void]$binding.Clear()
                $cleared++
            }
        }
    }

    return $cleared
}

function Get-WordKeyCode {
    param(
        [Parameter(Mandatory = $true)]
        $Word,

        [Parameter(Mandatory = $true)]
        [int[]]$Keys
    )

    switch ($Keys.Count) {
        1 {
            $key1 = $Keys[0]
            return $Word.BuildKeyCode($key1)
        }
        2 {
            $key1 = $Keys[0]
            $key2 = $Keys[1]
            return $Word.BuildKeyCode($key1, [ref]$key2)
        }
        3 {
            $key1 = $Keys[0]
            $key2 = $Keys[1]
            $key3 = $Keys[2]
            return $Word.BuildKeyCode($key1, [ref]$key2, [ref]$key3)
        }
        4 {
            $key1 = $Keys[0]
            $key2 = $Keys[1]
            $key3 = $Keys[2]
            $key4 = $Keys[3]
            return $Word.BuildKeyCode($key1, [ref]$key2, [ref]$key3, [ref]$key4)
        }
        default { throw "Unsupported keybinding shape: $($Keys -join ', ')" }
    }
}

function Install-WordifyKeyBindings {
    param(
        [Parameter(Mandatory = $true)]
        $Word,

        [Parameter(Mandatory = $true)]
        $Template,

        [Parameter(Mandatory = $true)]
        [object[]]$Bindings
    )

    $Word.CustomizationContext = $Template
    $installed = New-Object System.Collections.Generic.List[string]

    foreach ($binding in $Bindings) {
        $keyCode = Get-WordKeyCode -Word $Word -Keys $binding.Keys
        if ($PSCmdlet.ShouldProcess($binding.Name, "Assign to $($binding.Command)")) {
            [void]$Word.KeyBindings.Add($script:WdKeyCategoryMacro, $binding.Command, $keyCode)
            [void]$installed.Add($binding.Name)
        }
    }

    return $installed
}

function Save-WordTemplate {
    param(
        [Parameter(Mandatory = $true)]
        $Template
    )

    $target = [string]$Template.FullName
    if ($PSCmdlet.ShouldProcess($target, "Save Word template")) {
        [void]$Template.Save()
    }
}

$installerRoot = Get-InstallerPackageRoot -ScriptRoot $PSScriptRoot
$repoRoot = Get-BundledWordifySourceRoot -ScriptRoot $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($repoRoot)) {
    $repoRoot = $installerRoot
}

Assert-WordAutomationIsAvailable

if ((-not $Uninstall) -and [string]::IsNullOrWhiteSpace($VbaPath)) {
    if (Test-WordifySourceRoot -Root $repoRoot) {
        Write-Host "Using bundled Wordify source: $repoRoot"
    }
    elseif (-not $SkipSourceBootstrap) {
        Write-Host "No bundled Wordify source found. Starting from-scratch bootstrap."
        $bootstrapClonePath = Resolve-BootstrapClonePath `
            -PreferredPath $ClonePath `
            -InstallerRoot $installerRoot `
            -ClonePathWasExplicit $PSBoundParameters.ContainsKey("ClonePath")

        Write-Host "Wordify source checkout path: $bootstrapClonePath"
        $repoRoot = Initialize-WordifyInstallSource -RepositoryUrl $RepositoryUrl -ClonePath $bootstrapClonePath
    }
    else {
        throw @"
-SkipSourceBootstrap was requested, but no bundled Wordify source was found next to this installer.

Expected to find a 'vba' folder and 'manifest.xml' beside the installer package.
Run the installer without -SkipSourceBootstrap to clone the source code, or pass -VbaPath to an existing Wordify vba folder.
"@
    }
}

Assert-NormalTemplateIsAvailable

if ([string]::IsNullOrWhiteSpace($VbaPath)) {
    $VbaPath = Join-Path $repoRoot "vba"
}

$modules = $null
if (-not $Uninstall) {
    $modules = Resolve-WordifyModuleFiles -Root $VbaPath -Specs $ModuleSpecs
}

$word = $null
$bootstrapDocument = $null
try {
    $word = New-WordApplication -ShowWindow ([bool]$Visible)
    $bootstrapDocument = Initialize-WordVbaProjectModel -Word $word
    $template = $word.NormalTemplate
    $templatePath = [string]$template.FullName

    Write-Host "Using Word template: $templatePath"
    Assert-VbaProjectAccess -Template $template

    if (-not $SkipKeybindings) {
        $cleared = Clear-WordifyKeyBindings -Word $word -Template $template -Bindings $KeyBindings
        Write-Host "Cleared Wordify keybindings: $cleared"
    }

    $removed = @(Remove-WordifyModules -Template $template -Specs $ModuleSpecs)
    Write-Host "Removed Wordify modules: $($removed.Count)"

    if ($Uninstall) {
        Save-WordTemplate -Template $template
        $removedSourceFolders = @(Remove-WordifySourceFolders `
            -ClonePath $ClonePath `
            -ClonePathWasExplicit $PSBoundParameters.ContainsKey("ClonePath"))

        Write-Host "Removed Wordify source folders: $($removedSourceFolders.Count)"
        Write-Host "Wordify has been uninstalled from Normal.dotm."
        exit 0
    }

    $imported = @(Import-WordifyModules -Template $template -Modules $modules)
    Write-Host "Imported Wordify modules: $($imported.Count)"

    if (-not $SkipKeybindings) {
        $installed = @(Install-WordifyKeyBindings -Word $word -Template $template -Bindings $KeyBindings)
        Write-Host "Installed Wordify keybindings: $($installed -join ', ')"
    }

    Save-WordTemplate -Template $template
    Write-Host "Wordify has been installed into Normal.dotm."
}
finally {
    if ($null -ne $bootstrapDocument) {
        try {
            Close-WordDocumentNoSave -Document $bootstrapDocument
        }
        catch {
            Write-Warning "Could not close the temporary Word document cleanly: $($_.Exception.Message)"
        }

        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($bootstrapDocument)
    }

    if ($null -ne $word) {
        try {
            $saveChanges = $script:WdDoNotSaveChanges
            $missingOriginalFormat = [System.Type]::Missing
            $missingRouteDocument = [System.Type]::Missing
            [void]$word.Quit([ref]$saveChanges, [ref]$missingOriginalFormat, [ref]$missingRouteDocument)
        }
        catch {
            Write-Warning "Could not close the Word automation instance cleanly: $($_.Exception.Message)"
        }

        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
