#Requires -Version 5.1
#Requires -Modules @{ModuleName='PSWriteHTML';ModuleVersion='1.17.0'}

param(
    [System.IO.DirectoryInfo]$ReportPath,
    [System.IO.DirectoryInfo]$OutputPath = "$PSScriptRoot\output",
    [string]$DateFormat = 'yyyy-MM-dd',
    [string]$URI = 'https://www.purple-knight.com',
    [string]$Logo,
    [string]$Author = 'PurpleKnightDashboard',
    [int]$MaxWidth = 1400,
    [switch]$DoNotShow,
    [switch]$InvertChartLine
)

. "$PSScriptRoot\Functions.ps1"

$PSDefaultParameterValues = @{
    'New-HTMLSection:HeaderBackGroundColor' = $Colors.Neutral
    'New-HTMLSection:HeaderTextSize'        = 16
    'New-HTMLSection:Margin'                = 20
    'New-ChartBar:Color'                    = $Colors.Primary
    'New-ChartLine:Color'                   = $Colors.Primary
    'New-HTMLTable:HTML'                    = { { New-SeverityConditions } }
    'New-HTMLTable*:WarningAction'          = 'SilentlyContinue'
    'New-HTMLGage:MinValue'                 = 0
    'New-HTMLGage:MaxValue'                 = 100
    'New-HTMLGage:Pointer'                  = $true
    'New-HTMLGage:LevelColors'              = $GageColors
}

if (!$ReportPath) {
    $htmlFiles = Get-File -Directory "$PSScriptRoot\reports" -Filter 'Purple Knight HTML report (*.html)|*.html' | ForEach-Object { Get-Item -Path $_ }
}
else {
    $htmlFiles = Get-ChildItem -Path $ReportPath -Filter '*.html' -Recurse
}

if (!(Test-Path -Path $OutputPath.FullName -PathType Container)) {
    $null = New-Item -Path $OutputPath.FullName -ItemType Directory
}

$allReports = Import-PurpleKnightReports -Files $htmlFiles -ExceptionsPath "$PSScriptRoot\data\exceptions.csv" -DateFormat $DateFormat
if (!$allReports) {
    Write-Warning 'No Purple Knight report found'
    return
}

# Add month property to reports
$allReports | Add-Member -MemberType NoteProperty -Name 'Month' -Value $null -Force
$allReports | ForEach-Object { $_.Month = Get-Date $_.Date -Format 'yyyy-MM' }

# Keep only the last report for each domain and month
$reports = @($allReports | Sort-Object -Property Date -Descending | Group-Object -Property Domain, Month | ForEach-Object { $_.Group | Select-Object -First 1 } | Sort-Object -Property Date)

$domains = @($reports.Domain | Sort-Object -Unique)
$months = @($reports.Month | Sort-Object -Unique)
$allExposures = $reports.Exposures | Sort-Object -Unique -Property IndicatorId
$categoryNames = @($reports.Categories | Where-Object { $_.Included } | Sort-Object ID | ForEach-Object { $_.Name } | Select-Object -Unique)

# Limit to 10 domains for better readability of charts
if ($domains.Count -gt 10) {
    Write-Warning "More than 10 domains found, only the 10 first will be kept for charts readability"
    $domains = $domains | Select-Object -First 10
}

# Return one value per month for a domain ($null when there is no report this month)
function Get-MonthlyValues {
    param (
        [string]$Domain,
        [scriptblock]$Value
    )

    foreach ($month in $months) {
        $report = $reports | Where-Object { $_.Domain -eq $Domain -and $_.Month -eq $month }
        if ($report) { & $Value $report } else { $null }
    }
}

$LogoHtml = Get-LogoHtml -Logo $Logo -URI $URI
$filePath = "$OutputPath\dashboard_global.html"

New-HTML -Name 'Global - Purple Knight dashboard' -FilePath $filePath -Encoding UTF8 -Author $Author -DateFormat 'dd/MM/yyyy HH:mm:ss' {

    # Header
    New-HTMLHeader -HTMLContent {
        $domain = 'Global report'
        $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path "$PSScriptRoot\data\header.html" -Encoding UTF8))
    }

    # Main
    New-HTMLMain {

        New-HTMLTab -Name 'Home' -IconSolid home {

            $lastReports = $domains | ForEach-Object { $domain = $_ ; $reports | Where-Object { $_.Domain -eq $domain } | Select-Object -Last 1 }
            $lastIds = @($lastReports.Exposures.IndicatorId | Sort-Object -Unique)

            $riskSolvedSince = $allExposures | Where-Object { $_.IndicatorId -notin $lastIds } | ForEach-Object {
                $indicatorId = $_.IndicatorId
                $lastReport = @($reports | Where-Object { $_.Exposures.IndicatorId -contains $indicatorId })[-1]
                $_ | Select-Object Severity, ANSSI, Category, IndicatorId, Name,
                @{Name = 'LastAppearance'; Expression = { $lastReport.Month } },
                @{Name = 'LastDomain'; Expression = { $lastReport.Domain } }
            }

            # Exposed in the latest report of a domain but not in its first report
            $riskNewSince = $lastReports | ForEach-Object {
                $report = $_
                $domainReports = @($reports | Where-Object { $_.Domain -eq $report.Domain })
                $firstIds = @($domainReports[0].Exposures.IndicatorId)
                $report.Exposures | Where-Object { $_.IndicatorId -notin $firstIds } | ForEach-Object {
                    $indicatorId = $_.IndicatorId
                    $firstAppearance = @($domainReports | Where-Object { $_.Exposures.IndicatorId -contains $indicatorId })[0].Month
                    $_ | Select-Object @{Name = 'Domain'; Expression = { $report.Domain } }, Points, Severity, ANSSI, Category, IndicatorId, Name,
                    @{Name = 'FirstAppearance'; Expression = { $firstAppearance } }
                }
            }

            $latestScores = $lastReports | ForEach-Object {
                $report = $_
                $row = [ordered]@{
                    Domain           = $report.Domain
                    Date             = $report.Label
                    Grade            = $report.Grade
                    'Security score' = $report.Score
                    'ANSSI'          = $report.Maturity
                    'Risk points'    = Get-RiskPoints $report.Exposures
                    Exposures        = $report.Exposures.Count
                }
                $categoryNames | ForEach-Object {
                    $name = $_
                    $row[$name] = ($report.Categories | Where-Object { $_.Name -eq $name }).Score
                }
                [PSCustomObject]$row
            }

            $exposuresPerDomain = $allExposures | Select-Object Severity, ANSSI, Category, IndicatorId, Name
            $domains | ForEach-Object { $exposuresPerDomain | Add-Member -Name $_ -MemberType NoteProperty -Value $null }
            $lastReports | ForEach-Object {
                $report = $_
                $exposuresPerDomain | ForEach-Object {
                    $indicatorId = $_.IndicatorId
                    $_.($report.Domain) = ($report.Exposures | Where-Object { $_.IndicatorId -eq $indicatorId }).Points
                }
            }
            $exposuresPerDomain | Add-Member -Name 'Global' -MemberType NoteProperty -Value $null
            $exposuresPerDomain | ForEach-Object {
                $exposure = $_
                $global = $domains | ForEach-Object { $exposure.$_ } | Where-Object { $null -ne $_ }
                if (($global | Measure-Object).Count -ge 1) { $_.Global = ($global | Measure-Object -Sum).Sum }
            }

            # New exposures and remediations
            New-HTMLSection -HeaderText 'Improvement & deterioration (first to latest report of each domain)' {
                New-HTMLPanel {
                    New-HTMLSection -Invisible -Margin 0 -AlignItems center -JustifyContent flex-start -BackgroundColor $Colors.Negative {
                        New-HTMLHeading h2 -HeadingText "New exposures ($(@($riskNewSince).Count))"
                    }
                    New-HTMLTable -DataTable $riskNewSince -DefaultSortColumn 'Points' -DefaultSortOrder Descending -HideButtons -DisablePaging
                }
                New-HTMLPanel {
                    New-HTMLSection -Invisible -Margin 0 -AlignItems center -JustifyContent flex-start -BackgroundColor $Colors.Positive {
                        New-HTMLHeading h2 -HeadingText "Exposures resolved ($(@($riskSolvedSince).Count))"
                    }
                    New-HTMLTable -DataTable $riskSolvedSince -DefaultSortColumn 'LastAppearance' -DefaultSortOrder Descending -HideButtons -DisablePaging
                }
            }

            # Diagrams for global score
            New-HTMLSection -HeaderText 'Evolution of security score and exposures' {
                New-HTMLPanel {
                    New-HTMLSection -Invisible {
                        New-HTMLChart -Title 'Purple Knight security score (100 is the best)' {
                            New-ChartAxisX -Name $months
                            $i = 0
                            $domains | ForEach-Object {
                                New-ChartLine -Value (Get-MonthlyValues -Domain $_ -Value { param($r) $r.Score }) -Name $_ -Color $Palette[$i]
                                $i++
                            }
                        }
                        New-HTMLChart -Title 'Total risk points' {
                            New-ChartAxisX -Name $months
                            if ($InvertChartLine.IsPresent) { New-ChartAxisY -Show -Reversed }
                            $i = 0
                            $domains | ForEach-Object {
                                New-ChartLine -Value (Get-MonthlyValues -Domain $_ -Value { param($r) Get-RiskPoints $r.Exposures }) -Name $_ -Color $Palette[$i]
                                $i++
                            }
                        }
                    }
                    New-HTMLSection -Invisible {
                        New-HTMLChart -Title 'ANSSI maturity level' {
                            New-ChartAxisX -Name $months
                            New-ChartAxisY -Show -MinValue 1 -MaxValue 5
                            $i = 0
                            $domains | ForEach-Object {
                                New-ChartLine -Value (Get-MonthlyValues -Domain $_ -Value { param($r) $r.Maturity }) -Name $_ -Curve stepline -Color $Palette[$i]
                                $i++
                            }
                        }
                        New-HTMLChart -Title 'Exposures' {
                            New-ChartAxisX -Name $months
                            if ($InvertChartLine.IsPresent) { New-ChartAxisY -Show -Reversed }
                            $i = 0
                            $domains | ForEach-Object {
                                New-ChartLine -Value (Get-MonthlyValues -Domain $_ -Value { param($r) $r.Exposures.Count }) -Name $_ -Color $Palette[$i]
                                $i++
                            }
                        }
                    }
                }
            }

            # Diagrams per category (three per row)
            New-HTMLSection -HeaderText 'Evolution of the score per category' -Direction column {
                for ($c = 0; $c -lt $categoryNames.Count; $c += 3) {
                    New-HTMLSection -Invisible {
                        $categoryNames[$c..([math]::Min($c + 2, $categoryNames.Count - 1))] | ForEach-Object {
                            $name = $_
                            New-HTMLPanel {
                                New-HTMLChart -Title $name {
                                    New-ChartAxisX -Name $months
                                    $i = 0
                                    $domains | ForEach-Object {
                                        New-ChartLine -Value (Get-MonthlyValues -Domain $_ -Value { param($r) ($r.Categories | Where-Object { $_.Name -eq $name }).Score }) -Name $_ -Color $Palette[$i]
                                        $i++
                                    }
                                }
                            }
                        }
                    }
                }
            }

            # Latest scores
            New-HTMLSection -HeaderText 'Latest score per domain' {
                New-HTMLTable -DataTable $latestScores -DefaultSortIndex 0 -DisablePaging {
                    1..5 | ForEach-Object {
                        New-HTMLTableCondition -Name 'ANSSI' -ComparisonType number -Operator eq -Value $_ -BackgroundColor $Colors."Level$_"
                    }
                    New-ScoreConditions -Names (@('Security score') + $categoryNames)
                }
            }

            # Exposures per domain
            New-HTMLSection -HeaderText 'Exposures per domain (risk points in the latest report)' {
                New-HTMLTable -DataTable $exposuresPerDomain -DefaultSortColumn 'Global' -DefaultSortOrder Descending -DisablePaging {
                    New-SeverityConditions
                    New-HTMLTableCondition -Name 'Global' -ComparisonType string -Operator eq -Value '' -BackgroundColor 'lightgray'
                    $domains | ForEach-Object { New-HTMLTableCondition -Name $_ -ComparisonType string -Operator eq -Value '' -BackgroundColor 'lightgray' }
                    New-HTMLTableCondition -Name 'Global' -ComparisonType string -Operator eq -Value '' -Color 'darkgray' -Row
                }
            }
        }

        # Create a new tab for each domain with its latest report
        $domains | ForEach-Object {

            $domain = $_
            $domainReports = @($allReports | Where-Object { $_.Domain -eq $domain })
            $currentReport = $domainReports[-1]
            $previousReport = if ($domainReports.Count -gt 1) { $domainReports[-2] } else { $null }
            $initialReport = if ($domainReports.Count -gt 2) { $domainReports[0] } else { $null }

            New-HTMLTab -Name $domain -IconSolid globe {
                New-ReportContent -CurrentReport $currentReport -PreviousReport $previousReport -InitialReport $initialReport -Title 'Latest report and environment information'
            }
        }
    }

    # Footer
    New-HTMLFooter -HTMLContent { $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path "$PSScriptRoot\data\footer.html" -Encoding UTF8)) }
}

Set-DashboardWidth -Path $filePath -MaxWidth $MaxWidth
if (!$DoNotShow.IsPresent) { Start-Process $filePath }
