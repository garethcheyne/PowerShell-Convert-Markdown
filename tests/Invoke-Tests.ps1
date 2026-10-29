<#
.SYNOPSIS
    Runs the ConvertMarkdown test suite.

.DESCRIPTION
    By default runs only the fast tests: pure string-to-string checks over the
    parsing and rendering functions, no Word and no network. That is the set CI
    should gate on.

    -IncludeWord adds the end-to-end conversions, which automate Microsoft Word
    and take roughly a minute.

.PARAMETER IncludeWord
    Also run the tests tagged Word. Needs Microsoft Word installed.

.PARAMETER WordOnly
    Run only the Word-tagged tests.

.PARAMETER CI
    Emit NUnit XML to tests/results.xml and exit non-zero on failure.

.EXAMPLE
    .\tests\Invoke-Tests.ps1

.EXAMPLE
    .\tests\Invoke-Tests.ps1 -IncludeWord
#>
[CmdletBinding()]
param(
    [switch] $IncludeWord,
    [switch] $WordOnly,
    [switch] $CI
)

$ErrorActionPreference = 'Stop'

$pester = Get-Module -ListAvailable Pester |
    Where-Object { $_.Version.Major -ge 5 } |
    Sort-Object Version -Descending |
    Select-Object -First 1

if (-not $pester) {
    throw "Pester 5 or newer is required. Install it with:`n" +
    '    Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser -Force -SkipPublisherCheck'
}
Import-Module $pester.Path -Force

$config = New-PesterConfiguration
$config.Run.Path = $PSScriptRoot
$config.Run.PassThru = $true
$config.Output.Verbosity = 'Detailed'

if ($WordOnly) { $config.Filter.Tag = 'Word' }
elseif (-not $IncludeWord) { $config.Filter.ExcludeTag = 'Word' }

if ($CI) {
    $config.TestResult.Enabled = $true
    $config.TestResult.OutputPath = Join-Path $PSScriptRoot 'results.xml'
}

$result = Invoke-Pester -Configuration $config

Write-Host ''
Write-Host ("Passed {0}   Failed {1}   Skipped {2}   {3:N1}s" -f
    $result.PassedCount, $result.FailedCount, $result.SkippedCount, $result.Duration.TotalSeconds)

if (-not $IncludeWord -and -not $WordOnly) {
    Write-Host 'Word-tagged end-to-end tests were skipped. Add -IncludeWord to run them.' -ForegroundColor DarkGray
}

if ($CI -and $result.FailedCount -gt 0) { exit 1 }
