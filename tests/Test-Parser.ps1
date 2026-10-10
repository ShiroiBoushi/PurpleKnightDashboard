# Self-check of the report parser, no module needed: .\tests\Test-Parser.ps1
# Stops at the first wrong value, prints "All checks passed" otherwise

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\..\Functions.ps1"

function Assert-Equal {
    param ($Actual, $Expected, [string]$What)

    if ("$Actual" -ne "$Expected") { throw "FAIL: $What - expected '$Expected', got '$Actual'" }
    Write-Host "OK: $What"
}

$fixture = Get-Item "$PSScriptRoot\fixture.html"
$report = Import-PurpleKnightReport -Path $fixture

Assert-Equal $report.Domain 'contoso.com' 'forest name'
Assert-Equal $report.Score 72 'security score'
Assert-Equal $report.Grade 'C-' 'grade conversion'
Assert-Equal $report.ExecutionTime '00:01:02' 'execution time without fraction'
Assert-Equal $report.Categories.Count 2 'categories'
Assert-Equal $report.Exposures.Count 2 'exposures (status Failed)'
Assert-Equal $report.FailedToRun.Count 1 'indicators failed to run (status Error)'

$critical = $report.Exposures | Where-Object { $_.IndicatorId -eq 'SI000003' }
Assert-Equal $critical.ANSSI 1 'indicator in ANSSI levels 1 and 3 keeps the lowest'
Assert-Equal $critical.Points 80 'risk points = weight x (100 - score) / 10'
Assert-Equal $critical.MITRE 'Credential Access, Privilege Escalation' 'MITRE tactics'
Assert-Equal $report.Maturity 1 'ANSSI maturity = lowest exposed level'

$exceptions = Join-Path ([System.IO.Path]::GetTempPath()) "pk-exceptions-$PID.csv"
try {
    Set-Content -Path $exceptions -Value "Domain;IndicatorId`r`ncontoso.com;SI000003`r`nfabrikam.com;SI000002" -Encoding UTF8
    $filtered = Import-PurpleKnightReports -Files $fixture -ExceptionsPath $exceptions -DateFormat 'yyyy-MM-dd'
}
finally {
    Remove-Item -Path $exceptions -ErrorAction SilentlyContinue
}
Assert-Equal $filtered.Exposures.IndicatorId 'SI000002' 'exception removes the exposure of this forest only'
Assert-Equal $filtered.IgnoredExposures.IndicatorId 'SI000003' 'ignored exposure is kept apart'
Assert-Equal $filtered.Maturity 3 'ANSSI maturity recomputed after exceptions'
Assert-Equal $filtered.Label (Get-Date $filtered.Date -Format 'yyyy-MM-dd') 'report label'

$notReport = Import-PurpleKnightReport -Path (Get-Item $PSCommandPath) -WarningAction SilentlyContinue
Assert-Equal $notReport $null 'a file which is not a report is skipped'

Write-Host 'All checks passed' -ForegroundColor Green
