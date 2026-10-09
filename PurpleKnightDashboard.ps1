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
    $htmlFiles = Get-File -Directory $PSScriptRoot -Filter 'Purple Knight HTML report (*.html)|*.html' | ForEach-Object { Get-Item -Path $_ }
}
else {
    $htmlFiles = Get-ChildItem -Path $ReportPath -Filter '*.html' -Recurse
}

if (!(Test-Path -Path $OutputPath.FullName -PathType Container)) {
    $null = New-Item -Path $OutputPath.FullName -ItemType Directory
}

$reports = Import-PurpleKnightReports -Files $htmlFiles -ExceptionsPath "$PSScriptRoot\data\exceptions.csv" -DateFormat $DateFormat
if (!$reports) {
    Write-Warning 'No Purple Knight report found'
    return
}

$LogoHtml = Get-LogoHtml -Logo $Logo -URI $URI

# Create one dashboard foreach domain
$reports.Domain | Sort-Object -Unique | ForEach-Object {

    $domain = $_
    $domainReports = @($reports | Where-Object { $_.Domain -eq $domain })
    $allExposures = $domainReports.Exposures | Sort-Object -Unique -Property IndicatorId
    $categoryNames = @($domainReports.Categories | Where-Object { $_.Included } | Sort-Object ID | ForEach-Object { $_.Name } | Select-Object -Unique)
    $filePath = "$OutputPath\dashboard_$($domain -replace '[\\/:*?"<>|]', '_').html"

    New-HTML -Name "$domain - Purple Knight dashboard" -FilePath $filePath -Encoding UTF8 -Author $Author -DateFormat 'dd/MM/yyyy HH:mm:ss' {

        # Header
        New-HTMLHeader -HTMLContent {
            $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path "$PSScriptRoot\data\header.html" -Encoding UTF8))
        }

        # Main
        New-HTMLMain {

            # Home tab
            New-HTMLTab -Name 'Home' -IconSolid home {

                $firstReport = $domainReports[0]
                $lastReport = $domainReports[-1]
                $lastIds = @($lastReport.Exposures.IndicatorId)
                $riskSolvedSince = $allExposures | Where-Object { $_.IndicatorId -notin $lastIds } | ForEach-Object {
                    $indicatorId = $_.IndicatorId
                    $lastAppearance = @($domainReports | Where-Object { $_.Exposures.IndicatorId -contains $indicatorId })[-1].Label
                    $_ | Select-Object Severity, ANSSI, Category, IndicatorId, Name, @{Name = 'LastAppearance'; Expression = { $lastAppearance } }
                }

                $scores = $domainReports | ForEach-Object {
                    $report = $_
                    $row = [ordered]@{
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

                $exposuresEvolution = $allExposures | Select-Object Severity, ANSSI, Category, IndicatorId, Name
                $domainReports | Sort-Object Date -Descending | ForEach-Object {
                    $exposuresEvolution | Add-Member -Name $_.Label -MemberType NoteProperty -Value $null
                }
                $domainReports | ForEach-Object {
                    $report = $_
                    $exposuresEvolution | ForEach-Object {
                        $indicatorId = $_.IndicatorId
                        $_.($report.Label) = ($report.Exposures | Where-Object { $_.IndicatorId -eq $indicatorId }).Points
                    }
                }

                $chartAxisX = $domainReports.Label
                $chartLineScore = $domainReports | ForEach-Object { $_.Score }
                $chartLinePoints = $domainReports | ForEach-Object { Get-RiskPoints $_.Exposures }
                $chartLineMaturity = $domainReports | ForEach-Object { [int]$_.Maturity }

                # Diagrams for global score
                New-HTMLSection -HeaderText 'Evolution of security score and exposures' {
                    New-HTMLPanel {
                        New-HTMLSection -Invisible {
                            New-HTMLChart -Title 'Purple Knight security score (100 is the best)' {
                                New-ChartAxisX -Name $chartAxisX
                                New-ChartLine -Value $chartLineScore -Name 'Score'
                            }
                            New-HTMLChart -Title 'Total risk points' {
                                New-ChartAxisX -Name $chartAxisX
                                if ($InvertChartLine.IsPresent) { New-ChartAxisY -Show -Reversed }
                                New-ChartLine -Value $chartLinePoints -Name 'Point(s)' -Color $Colors.Secondary
                            }
                        }
                        New-HTMLSection -Invisible {
                            if ($domainReports[0].HasAnssiLevels) {
                                New-HTMLChart -Title 'ANSSI maturity level' {
                                    New-ChartAxisX -Name $chartAxisX
                                    New-ChartAxisY -Show -MinValue 1 -MaxValue 5
                                    New-ChartLine -Value $chartLineMaturity -Curve stepline -Color $Colors.Level1 -Name 'Level'
                                }
                            }
                            New-HTMLChart -Title 'Exposures per severity' {
                                New-ChartAxisX -Name $chartAxisX
                                if ($InvertChartLine.IsPresent) { New-ChartAxisY -Show -Reversed }
                                $Severities | ForEach-Object {
                                    $severity = $_
                                    $values = $domainReports | ForEach-Object { @($_.Exposures | Where-Object { $_.Severity -eq $severity }).Count }
                                    New-ChartLine -Value $values -Name $severity -Color $Colors.$severity
                                }
                            }
                        }
                    }
                }

                # Diagrams per category (three per row)
                New-HTMLSection -HeaderText 'Evolution of the score per category' -Direction column {
                    for ($i = 0; $i -lt $categoryNames.Count; $i += 3) {
                        New-HTMLSection -Invisible {
                            $categoryNames[$i..([math]::Min($i + 2, $categoryNames.Count - 1))] | ForEach-Object {
                                $name = $_
                                New-HTMLPanel {
                                    New-HTMLChart -Title $name {
                                        New-ChartAxisX -Name $chartAxisX
                                        New-ChartLine -Value ($domainReports | ForEach-Object { [int]($_.Categories | Where-Object { $_.Name -eq $name }).Score }) -Name 'Score'
                                    }
                                }
                            }
                        }
                    }
                }

                # Remediations
                New-HTMLSection -HeaderText 'Remediations' {
                    New-HTMLTable -Title 'All exposures solved' -DataTable $riskSolvedSince -DefaultSortColumn 'LastAppearance' -DefaultSortOrder Descending -DisablePaging
                }

                # Scores
                New-HTMLSection -HeaderText 'Score & maturity evolution' {
                    New-HTMLTable -DataTable $scores -DefaultSortIndex 0 -DisablePaging {
                        1..5 | ForEach-Object {
                            New-HTMLTableCondition -Name 'ANSSI' -ComparisonType number -Operator eq -Value $_ -BackgroundColor $Colors."Level$_"
                        }
                        New-ScoreConditions -Names (@('Security score') + $categoryNames)
                    }
                }

                # Exposures evolution
                New-HTMLSection -HeaderText 'Exposures evolution (risk points)' {
                    New-HTMLTable -DataTable $exposuresEvolution -DefaultSortColumn $domainReports[-1].Label -DefaultSortOrder Descending -DisablePaging {
                        New-SeverityConditions
                        $domainReports.Label | ForEach-Object {
                            New-HTMLTableCondition -Name $_ -ComparisonType string -Operator eq -Value '' -BackgroundColor 'lightgray'
                        }
                        # Grayed the rows that have been resolved
                        New-HTMLTableCondition -Name $domainReports[-1].Label -ComparisonType string -Operator eq -Value '' -Color 'darkgray' -Row
                    }
                }
            }

            # Create a new tab for each report
            for ($i = 0; $i -lt $domainReports.Count; $i++) {
                $currentReport = $domainReports[$i]
                $previousReport = if ($i -gt 0) { $domainReports[$i - 1] } else { $null }
                $initialReport = if ($i -gt 1) { $domainReports[0] } else { $null }

                New-HTMLTab -Name $currentReport.Label {
                    New-ReportContent -CurrentReport $currentReport -PreviousReport $previousReport -InitialReport $initialReport
                }
            }
        }

        # Footer
        New-HTMLFooter -HTMLContent { $ExecutionContext.InvokeCommand.ExpandString([string](Get-Content -Path "$PSScriptRoot\data\footer.html" -Encoding UTF8)) }
    }

    Set-DashboardWidth -Path $filePath -MaxWidth $MaxWidth
    if (!$DoNotShow.IsPresent) { Start-Process $filePath }
}
