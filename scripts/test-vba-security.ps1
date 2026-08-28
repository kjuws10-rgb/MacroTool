[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$mainModulePath = Join-Path $repositoryRoot 'SecureExcelTransferVBA\SecureTransfer.bas'
$adminModulePath = Join-Path $repositoryRoot 'SecureExcelTransferVBA\SecureTransferAdminSetup.bas'

$failures = [System.Collections.Generic.List[string]]::new()

function Add-Failure {
    param([Parameter(Mandatory)][string]$Message)
    $script:failures.Add($Message)
}

function Assert-Matches {
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][string]$Pattern,
        [Parameter(Mandatory)][string]$Message
    )

    if (-not [regex]::IsMatch($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::Multiline)) {
        Add-Failure $Message
    }
}

function Assert-NotMatches {
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][string]$Pattern,
        [Parameter(Mandatory)][string]$Message
    )

    if ([regex]::IsMatch($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::Multiline)) {
        Add-Failure $Message
    }
}

function Get-CanonicalKeys {
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][string]$FunctionName
    )

    $functionPattern = "(?ms)^Private Function $([regex]::Escape($FunctionName))\b.*?^End Function\s*$"
    $functionMatch = [regex]::Match($Text, $functionPattern)
    if (-not $functionMatch.Success) {
        Add-Failure "Canonical configuration function not found: $FunctionName"
        return @()
    }

    return @([regex]::Matches($functionMatch.Value, '"([A-Za-z][A-Za-z0-9]+)="') |
        ForEach-Object { $_.Groups[1].Value })
}

function Assert-VbaBlockBalance {
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][string]$Label
    )

    $subOpen = [regex]::Matches($Text, '(?im)^\s*(?:Public\s+|Private\s+)?Sub\s+[A-Za-z_]').Count
    $subClose = [regex]::Matches($Text, '(?im)^\s*End Sub\s*$').Count
    $functionOpen = [regex]::Matches($Text, '(?im)^\s*(?:Public\s+|Private\s+)?Function\s+[A-Za-z_]').Count
    $functionClose = [regex]::Matches($Text, '(?im)^\s*End Function\s*$').Count

    if ($subOpen -ne $subClose) {
        Add-Failure "$Label Sub block count mismatch: $subOpen / $subClose"
    }
    if ($functionOpen -ne $functionClose) {
        Add-Failure "$Label Function block count mismatch: $functionOpen / $functionClose"
    }
}

function Assert-VbaProcedureNamesUnique {
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][string]$Label
    )

    $names = @([regex]::Matches(
        $Text,
        '(?im)^\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)'
    ) | ForEach-Object { $_.Groups[1].Value.ToLowerInvariant() })

    foreach ($duplicate in @($names | Group-Object | Where-Object Count -gt 1)) {
        Add-Failure "$Label duplicate procedure name: $($duplicate.Name)"
    }
}

function Assert-VbaContinuationLimit {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string[]]$Lines,
        [Parameter(Mandatory)][string]$Label
    )

    $continuations = 0
    for ($index = 0; $index -lt $Lines.Count; $index++) {
        if ($Lines[$index] -match '_\s*$') {
            $continuations++
            if ($continuations -gt 24) {
                Add-Failure "$Label line $($index + 1) exceeds 24 consecutive VBA line continuations."
            }
        }
        else {
            $continuations = 0
        }
    }
}

if (-not (Test-Path -LiteralPath $mainModulePath) -or -not (Test-Path -LiteralPath $adminModulePath)) {
    throw 'The VBA modules to test were not found.'
}

$mainModule = Get-Content -LiteralPath $mainModulePath -Raw
$adminModule = Get-Content -LiteralPath $adminModulePath -Raw
$allModules = $mainModule + "`n" + $adminModule

Assert-VbaBlockBalance -Text $mainModule -Label 'SecureTransfer.bas'
Assert-VbaBlockBalance -Text $adminModule -Label 'SecureTransferAdminSetup.bas'
Assert-VbaProcedureNamesUnique -Text $mainModule -Label 'SecureTransfer.bas'
Assert-VbaProcedureNamesUnique -Text $adminModule -Label 'SecureTransferAdminSetup.bas'

Assert-Matches $mainModule 'Private Const CONFIG_SCHEMA_VERSION As String = "2"' 'Runtime configuration schema is not version 2.'
Assert-Matches $adminModule 'Private Const ADMIN_SCHEMA_VERSION As String = "2"' 'Admin configuration schema is not version 2.'

$expectedKeys = @(
    'SchemaVersion',
    'DeploymentState',
    'ApprovalStatus',
    'AllowedRoot',
    'ApprovedWorkbookFullName',
    'ApprovedBy',
    'ApprovalExpiry',
    'ExportApprovalStatus',
    'AllowedExportRoot',
    'AllowedExportFormats',
    'MaxRows',
    'MaxColumns',
    'MaxCells',
    'MaxFileMB',
    'CsvCharset'
)

$mainCanonicalKeys = Get-CanonicalKeys -Text $mainModule -FunctionName 'CanonicalConfig'
$adminCanonicalKeys = Get-CanonicalKeys -Text $adminModule -FunctionName 'AdminCanonicalConfig'
if (($mainCanonicalKeys -join '|') -ne ($expectedKeys -join '|')) {
    Add-Failure "Runtime canonical key order differs: $($mainCanonicalKeys -join ', ')"
}
if (($adminCanonicalKeys -join '|') -ne ($expectedKeys -join '|')) {
    Add-Failure "Admin canonical key order differs: $($adminCanonicalKeys -join ', ')"
}

foreach ($key in $expectedKeys) {
    $escapedKey = [regex]::Escape($key)
    Assert-Matches $mainModule ('RequireSetting cfg, "{0}"' -f $escapedKey) "Runtime required-setting check missing: $key"
    Assert-Matches $adminModule ('cfg\.Add "{0}"' -f $escapedKey) "Admin setting creation missing: $key"
}
Assert-Matches $mainModule 'RequireSetting cfg, "ConfigChecksum"' 'Runtime checksum requirement is missing.'
Assert-Matches $adminModule 'cfg\.Add "ConfigChecksum"' 'Admin checksum creation is missing.'

Assert-Matches $mainModule 'Content-Security-Policy' 'HTML Content Security Policy is missing.'
Assert-Matches $mainModule "default-src 'none'" 'HTML default external-resource denial is missing.'
Assert-Matches $mainModule 'CreateTextFile\(outputPath, False, False\)' 'No-overwrite output reservation is missing.'
Assert-Matches $mainModule 'DeleteFile outputPath, True' 'Failed-output cleanup is missing.'
Assert-Matches $mainModule 'Application\.Workbooks\.Add\(xlWBATWorksheet\)' 'In-memory PDF value workbook is missing.'
Assert-Matches $mainModule 'temporaryRange\.NumberFormat = "@"' 'PDF text-only range formatting is missing.'
Assert-Matches $mainModule 'If RangeContainsFormula\(temporaryRange\)' 'PDF formula verification is missing.'
Assert-Matches $mainModule 'If fileSystem\.FileExists\(normalizedOutput\)' 'Existing output overwrite protection is missing.'
Assert-Matches $mainModule 'ExportApprovalStatus' 'Export approval validation is missing.'
Assert-Matches $mainModule 'AllowedExportRoot' 'Approved export root validation is missing.'
Assert-Matches $mainModule 'AllowedExportFormats' 'Allowed export format validation is missing.'
Assert-Matches $mainModule 'Public Sub SecureTransfer_ExportSelectedRangeHtml\(\)' 'Public HTML export entry point is missing.'
Assert-Matches $mainModule 'Public Sub SecureTransfer_ExportSelectedRangePdf\(\)' 'Public PDF export entry point is missing.'
Assert-Matches $adminModule '"btnExportHtml"' 'HTML home-screen button is missing.'
Assert-Matches $adminModule '"btnExportPdf"' 'PDF home-screen button is missing.'
Assert-Matches $mainModule 'Dim recordValues\(1 To 1, 1 To 4\) As Variant' 'Audit record is not limited to four fields.'
Assert-Matches $mainModule '(?s)AppendAuditRecord\(auditSheet,\s*_\s*ThisWorkbook\.Name,\s*_\s*FileNameOnly\(outputPath\)' 'Export audit does not use file names only.'

Assert-NotMatches $allModules '(?im)^\s*(?!'').*\bApplication\.SendKeys\b' 'SendKeys usage detected.'
Assert-NotMatches $allModules '(?im)^\s*(?!'').*CreateObject\("WScript\.Shell"\)' 'WScript.Shell usage detected.'
Assert-NotMatches $allModules '(?im)^\s*(?!'').*\bDeclare\s+(?:PtrSafe\s+)?(?:Function|Sub)\b' 'External API declaration detected.'
Assert-NotMatches $allModules '(?im)^\s*(?!'').*\bShell\s*\(' 'Process launch detected.'
Assert-NotMatches $allModules '(?im)^\s*(?!'').*FollowHyperlink' 'Hyperlink launch detected.'
Assert-NotMatches $mainModule '<script\b' 'Generated HTML includes a script tag.'

$mainLines = Get-Content -LiteralPath $mainModulePath
$adminLines = Get-Content -LiteralPath $adminModulePath
foreach ($lineInfo in @(
    @{ Label = 'SecureTransfer.bas'; Lines = $mainLines },
    @{ Label = 'SecureTransferAdminSetup.bas'; Lines = $adminLines }
)) {
    Assert-VbaContinuationLimit -Lines $lineInfo.Lines -Label $lineInfo.Label
    for ($index = 0; $index -lt $lineInfo.Lines.Count; $index++) {
        if ($lineInfo.Lines[$index].Length -gt 1023) {
            Add-Failure "$($lineInfo.Label) line $($index + 1) exceeds the safe VBA line length."
        }
    }
}

if ($failures.Count -gt 0) {
    Write-Host 'VBA security static checks failed:' -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    exit 1
}

Write-Host 'VBA security static checks passed: schema, export boundaries, HTML/PDF safeguards, and forbidden APIs.' -ForegroundColor Green
exit 0
