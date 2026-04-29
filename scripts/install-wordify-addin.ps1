[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
param(
    [Parameter(Mandatory = $false)]
    [string]$VbaPath,

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

Assert-WordAutomationIsAvailable
Assert-NormalTemplateIsAvailable

$repoRoot = Split-Path -Parent $PSScriptRoot
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
