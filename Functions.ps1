# Shared code for PurpleKnightDashboard.ps1 and PurpleKnightGlobalDashboard.ps1

$Colors = @{
    'Primary'       = '#6A2C91'
    'Secondary'     = '#3845AB'
    'Neutral'       = '#3D3834'
    'Positive'      = '#CFE9CF'
    'Negative'      = '#FFCECE'
    'Critical'      = '#F94144'
    'High'          = '#F8961E'
    'Medium'        = '#F9C74F'
    'Low'           = '#43AA8B'
    'Informational' = '#277DA1'
    'Level1'        = '#F94144'
    'Level2'        = '#F8961E'
    'Level3'        = '#F9C74F'
    'Level4'        = '#43AA8B'
    'Level5'        = '#277DA1'
    'Score1'        = 'darkred'
    'Score2'        = 'darkorange'
    'Score3'        = 'darkgoldenrod'
    'Score4'        = 'darkgreen'
    'Score5'        = 'darkcyan'
}

$Palette = @(
    '#F94144',
    '#90BE6D',
    '#256EFF',
    '#F9C74F',
    '#DA659A',
    '#43AA8B',
    '#1e93c5',
    '#8338EC',
    '#A9DEF9',
    '#F9844A'
)

$Severities = 'Critical', 'High', 'Medium', 'Low', 'Informational'

# Purple Knight scores go from 0 (worst, red) to 100 (best, green)
$GageColors = @('#ff0000', '#f9c802', '#a9d70b')

$Svg = @{
    Up    = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512" style="height: 25px; fill: {0};"><path d="M384 160c-17.7 0-32-14.3-32-32s14.3-32 32-32H544c17.7 0 32 14.3 32 32V288c0 17.7-14.3 32-32 32s-32-14.3-32-32V205.3L342.6 374.6c-12.5 12.5-32.8 12.5-45.3 0L192 269.3 54.6 406.6c-12.5 12.5-32.8 12.5-45.3 0s-12.5-32.8 0-45.3l160-160c12.5-12.5 32.8-12.5 45.3 0L320 306.7 466.7 160H384z"/></svg>'
    Down  = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512" style="height: 25px; fill: {0};"><path d="M384 352c-17.7 0-32 14.3-32 32s14.3 32 32 32H544c17.7 0 32-14.3 32-32V224c0-17.7-14.3-32-32-32s-32 14.3-32 32v82.7L342.6 137.4c-12.5-12.5-32.8-12.5-45.3 0L192 242.7 54.6 105.4c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3l160 160c12.5 12.5 32.8 12.5 45.3 0L320 205.3 466.7 352H384z"/></svg>'
    Equal = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" style="height: 22px; fill: {0};"><path d="M502.6 278.6c12.5-12.5 12.5-32.8 0-45.3l-128-128c-12.5-12.5-32.8-12.5-45.3 0s-12.5 32.8 0 45.3L402.7 224 32 224c-17.7 0-32 14.3-32 32s14.3 32 32 32l370.7 0-73.4 73.4c-12.5 12.5-12.5 32.8 0 45.3s32.8 12.5 45.3 0l128-128z"/></svg>'
}

function Get-File {
    param (
        [string]$Directory = 'C:\',
        [string]$Filter = 'All files (*.*)|*.*'
    )

    $null = [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = (Get-Item $Directory).FullName
    $OpenFileDialog.Filter = $Filter
    $OpenFileDialog.Multiselect = $true
    $null = $OpenFileDialog.ShowDialog()

    $OpenFileDialog.FileNames
}

function ConvertTo-LocalDate {
    param ($Value)

    # PowerShell 7 already converts ISO 8601 strings to [datetime], Windows PowerShell 5.1 does not
    if ($Value -isnot [datetime]) {
        $Value = [datetime]::Parse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
    }
    $Value.ToLocalTime()
}

function ConvertTo-Grade {
    param ([string]$Grade)

    $Grade -replace 'Plus$', '+' -replace 'Minus$', '-'
}

function Get-AnssiMaturity {
    param ($Report)

    # ANSSI maturity is the lowest level which still has an exposure, or 5 if every level is clean
    if (!$Report.HasAnssiLevels) { return $null }
    $levels = @($Report.Exposures | Where-Object { $_.ANSSI } | ForEach-Object { [int]$_.ANSSI })
    if ($levels.Count -gt 0) { ($levels | Measure-Object -Minimum).Minimum } else { 5 }
}

function Get-RiskPoints {
    param ($Exposures)

    [int]($Exposures | Measure-Object -Property Points -Sum).Sum
}

function Import-PurpleKnightReport {
    param (
        [Parameter(Mandatory)]
        [System.IO.FileInfo]$Path
    )

    $content = [System.IO.File]::ReadAllText($Path.FullName, [System.Text.Encoding]::UTF8)
    $scripts = [regex]::Matches($content, '<script>([\s\S]*?)</script>') | ForEach-Object { $_.Groups[1].Value }

    # Global report data: window.reportJSON = {...};
    $reportScript = $scripts | Where-Object { $_ -match '^\s*window\.reportJSON\s*=' } | Select-Object -First 1
    if (!$reportScript) {
        Write-Warning "'$($Path.Name)' is not a Purple Knight HTML report, skipped"
        return
    }
    $reportJson = ($reportScript -replace '^\s*window\.reportJSON\s*=\s*' -replace ';\s*$') | ConvertFrom-Json

    # Indicator data: window["Category_X"]["<uuid>"] = {...};
    $indicators = @{}
    $scripts | Where-Object { $_ -match 'window\["Category_\d+"\]\["' } | ForEach-Object {
        $code = $_
        $blocks = [regex]::Matches($code, 'window\["Category_(\d+)"\]\["([0-9a-fA-F-]+)"\]\s*=\s*')
        for ($i = 0; $i -lt $blocks.Count; $i++) {
            $start = $blocks[$i].Index + $blocks[$i].Length
            $end = if ($i + 1 -lt $blocks.Count) { $blocks[$i + 1].Index } else { $code.Length }
            $json = $code.Substring($start, $end - $start).Trim() -replace ';$'
            $indicators["$($blocks[$i].Groups[1].Value)/$($blocks[$i].Groups[2].Value)"] = $json | ConvertFrom-Json
        }
    }

    # ANSSI level of each indicator (only available for Active Directory)
    # An indicator can be listed in several levels, the lowest one is kept
    $anssiLevels = @{}
    $reportJson.ANSIIAppendix | Where-Object { $_ } | ForEach-Object {
        $level = [int]$_.Level
        $_.Indicators | ForEach-Object {
            if (!$anssiLevels.ContainsKey($_) -or $anssiLevels[$_] -gt $level) { $anssiLevels[$_] = $level }
        }
    }

    foreach ($result in $reportJson.reportResultsList) {

        $name = @($result.ForestName, $result.TenantName, $result.OktaDomainUrl) | Where-Object { $_ } | Select-Object -First 1
        $summary = $result.ReportResults

        $categories = foreach ($category in $result.ReportCategories) {
            [PSCustomObject]@{
                ID       = [int]$category.Category.ID
                Name     = $category.Category.Name
                Score    = [int]$category.Category.Score
                Grade    = ConvertTo-Grade $category.Category.Grade
                Weight   = [int]$category.Category.Weight
                Included = [bool]$category.Category.IncludeInCalculation
            }
        }

        $allIndicators = foreach ($category in $result.ReportCategories) {
            foreach ($id in $category.Indicators) {
                $indicator = $indicators["$($category.Category.ID)/$id"]
                if (!$indicator) { continue }
                $info = $indicator.ResIndicator
                $exec = $indicator.ExecutionResult
                $status = if ($exec) { $exec.Status } else { $info.State }
                $score = if ($exec) { [int]$exec.Score } else { $null }
                [PSCustomObject]@{
                    Points      = if ($info.IsFailed) { [int][math]::Round($info.Weight * (100 - $score) / 10) } else { 0 }
                    Severity    = $info.Severity
                    ANSSI       = $anssiLevels[$id]
                    Category    = $category.Category.Name
                    IndicatorId = $info.ShortName
                    Name        = $info.Name
                    Score       = $score
                    Results     = [int]$indicator.TotalResultsCount
                    Weight      = [int]$info.Weight
                    Exposure    = $info.ExposureType
                    Status      = $status
                    Message     = if ($exec) { $exec.ResultMessage } else { $null }
                    MITRE       = ($info.SecurityFrameworks | Where-Object { $_.Name -eq 'MITRE ATT&CK' }).Tags -join ', '
                    UUID        = $id
                }
            }
        }

        $report = [PSCustomObject]@{
            Domain           = $name
            Environment      = $result.ADType
            Domains          = $result.SelectedDomains -join ', '
            Date             = ConvertTo-LocalDate $summary.GeneratedOn
            Label            = $null
            Version          = $reportJson.AppVersion
            Edition          = $reportJson.Edition.Name
            RunAs            = $result.UserName
            ExecutionTime    = ([string]$summary.TestExecutionTime) -replace '\.\d+$'
            Score            = [int]$summary.TotalScore
            Grade            = ConvertTo-Grade $summary.TotalGrade
            Maturity         = $null
            HasAnssiLevels   = $anssiLevels.Count -gt 0 -and $result.ADType -eq 'AD'
            Evaluated        = [int]$summary.TotalEvaluatedCount
            Passed           = [int]$summary.PassedCount
            Categories       = $categories
            Indicators       = $allIndicators
            Exposures        = @($allIndicators | Where-Object { $_.Status -eq 'Failed' })
            FailedToRun      = @($allIndicators | Where-Object { $_.Status -eq 'Error' })
            IgnoredExposures = @()
            File             = $Path.Name
        }
        $report.Maturity = Get-AnssiMaturity $report
        $report
    }
}

function Import-PurpleKnightReports {
    param (
        [System.IO.FileInfo[]]$Files,
        [string]$ExceptionsPath,
        [string]$DateFormat
    )

    $reports = $Files | ForEach-Object {
        Write-Verbose "Importing $($_.FullName)"
        Import-PurpleKnightReport -Path $_
    }

    # Handle exceptions
    if (Test-Path -Path $ExceptionsPath -PathType Leaf) {
        $exceptions = Import-Csv -Path $ExceptionsPath -Delimiter ';' -Encoding UTF8
        $reports | ForEach-Object {
            $domain = $_.Domain
            $ignored = @(($exceptions | Where-Object { $_.Domain -eq $domain -or $_.Domain -eq '*' }).IndicatorId)
            $_.IgnoredExposures = @($_.Exposures | Where-Object { $_.IndicatorId -in $ignored -or $_.UUID -in $ignored })
            $_.Exposures = @($_.Exposures | Where-Object { $_.IndicatorId -notin $ignored -and $_.UUID -notin $ignored })
            $_.Maturity = Get-AnssiMaturity $_
        }
    }

    $reports = $reports | Sort-Object Date

    # Give each report a unique label per domain (used for tab names and table columns)
    $reports | Group-Object Domain | ForEach-Object {
        $_.Group | Group-Object { Get-Date $_.Date -Format $DateFormat } | ForEach-Object {
            $n = 0
            $_.Group | ForEach-Object {
                $n++
                $_.Label = if ($n -eq 1) { $_.Date.ToString($DateFormat) } else { "$($_.Date.ToString($DateFormat)) ($n)" }
            }
        }
    }

    $reports
}

function Get-TrendHtml {
    param (
        $Current,
        $Previous,
        [switch]$HigherIsBetter,
        [string]$Text
    )

    $arrow = $null
    if ($null -ne $Previous) {
        $good = $Colors.Low
        $bad = $Colors.Critical
        if ($HigherIsBetter) { $good, $bad = $bad, $good }
        if ($Previous -lt $Current) { $arrow = $Svg.Up -f $bad }
        elseif ($Previous -gt $Current) { $arrow = $Svg.Down -f $good }
        else { $arrow = $Svg.Equal -f 'gray' }
    }

    @'
<div style="justify-content:center;align-items: center;width: 100%;display: flex;gap: 8px;">
  {0}
  <div class="defaultText">
    <div align="center">
      <span style="font-weight:bold;text-align:center;font-size:22px">{1}</span>
    </div>
  </div>
</div>
'@ -f $arrow, $Text
}

function Get-BadgeHtml {
    param (
        [string]$Text,
        [string]$Color
    )

    '<span style="text-align:center;font-size:12px;color:#ffffff;padding: 2px 4px;background-color:{0};border-radius:2px;">{1}</span>' -f $Color, $Text
}

function New-ScoreConditions {
    param ([string[]]$Names)

    # Purple Knight scores go from 0 (worst) to 100 (best)
    $Names | ForEach-Object {
        New-HTMLTableCondition -Name $_ -ComparisonType number -Operator ge -Value 0 -Color $Colors.Score1 -FontWeight bold
        New-HTMLTableCondition -Name $_ -ComparisonType number -Operator ge -Value 60 -Color $Colors.Score2 -FontWeight bold
        New-HTMLTableCondition -Name $_ -ComparisonType number -Operator ge -Value 70 -Color $Colors.Score3 -FontWeight bold
        New-HTMLTableCondition -Name $_ -ComparisonType number -Operator ge -Value 80 -Color $Colors.Score4 -FontWeight bold
        New-HTMLTableCondition -Name $_ -ComparisonType number -Operator ge -Value 90 -Color $Colors.Score5 -FontWeight bold
    }
}

function New-SeverityConditions {
    $Severities | ForEach-Object {
        New-HTMLTableCondition -Name 'Severity' -ComparisonType string -Operator eq -Value $_ -BackgroundColor $Colors.$_
    }
    1..5 | ForEach-Object {
        New-HTMLTableCondition -Name 'ANSSI' -ComparisonType number -Operator eq -Value $_ -BackgroundColor $Colors."Level$_"
    }
}

function Select-ExposureColumns {
    param ($Exposures)

    $Exposures | Select-Object Points, Severity, ANSSI, Category, IndicatorId, Name, Score, Results, Message
}

function New-ReportContent {
    # Content of a report tab: one Purple Knight report compared to the previous and initial ones
    param (
        $CurrentReport,
        $PreviousReport,
        $InitialReport,
        [string]$Title = 'Report and environment information'
    )

    if ($PreviousReport) {
        $previousIds = @($PreviousReport.Exposures.IndicatorId)
        $currentIds = @($CurrentReport.Exposures.IndicatorId)
        $riskSolved = Select-ExposureColumns ($PreviousReport.Exposures | Where-Object { $_.IndicatorId -notin $currentIds })
        $riskNew = Select-ExposureColumns ($CurrentReport.Exposures | Where-Object { $_.IndicatorId -notin $previousIds })
    }
    else {
        $riskSolved = $null
        $riskNew = $null
    }

    $categoryNames = @($CurrentReport.Categories | Where-Object { $_.Included } | ForEach-Object { $_.Name })

    # Main information about the report and the environment
    New-HTMLSection -HeaderText $Title -Direction column {
        New-HTMLSection -Invisible {
            New-HTMLPanel {
                $mainInfo = [PSCustomObject]@{
                    'Purple Knight version' = "$($CurrentReport.Version) ($($CurrentReport.Edition))"
                    'Generated on'          = Get-Date $CurrentReport.Date -Format F
                    'Report age'            = "$([int]((New-TimeSpan -Start $CurrentReport.Date).TotalDays)) day(s)"
                    'Environment'           = $CurrentReport.Environment
                    'Scanned domain(s)'     = $CurrentReport.Domains
                    'ANSSI maturity level'  = $CurrentReport.Maturity
                    'Run as'                = $CurrentReport.RunAs
                    'Execution time'        = $CurrentReport.ExecutionTime
                    'Indicators'            = "$($CurrentReport.Evaluated) evaluated, $($CurrentReport.Passed) passed, $($CurrentReport.Exposures.Count) exposure(s), $($CurrentReport.FailedToRun.Count) failed to run"
                    'Source file'           = $CurrentReport.File
                }
                New-HTMLTable -Title 'Report information' -DataTable $mainInfo -HideFooter -Transpose -Simplify
            }
            New-HTMLPanel {
                New-HTMLGage -Label 'Security score' -Value $CurrentReport.Score
                New-HTMLText -Alignment center -FontSize 18 -FontWeight bold -Text "Grade $($CurrentReport.Grade)"
                New-HTMLText -Alignment center -Text 'Overall score computed by Purple Knight (100 is the best)'
            }
            New-HTMLPanel {
                New-HTMLChart -Title 'Exposures per severity' {
                    $Severities | ForEach-Object {
                        $severity = $_
                        New-ChartPie -Value @($CurrentReport.Exposures | Where-Object { $_.Severity -eq $severity }).Count -Name $severity -Color $Colors.$severity
                    }
                }
            }
        }

        # Risk points with trend compared to the previous report
        New-HTMLSection -Invisible {
            New-HTMLPanel {
                $current = Get-RiskPoints $CurrentReport.Exposures
                $previous = if ($PreviousReport) { Get-RiskPoints $PreviousReport.Exposures } else { $null }
                Get-TrendHtml -Current $current -Previous $previous -Text "$current pt(s)"
                New-HTMLText -Text 'Total risk points' -Alignment center -FontSize 12
            }
            $Severities | ForEach-Object {
                $severity = $_
                New-HTMLPanel {
                    $current = Get-RiskPoints ($CurrentReport.Exposures | Where-Object { $_.Severity -eq $severity })
                    $previous = if ($PreviousReport) { Get-RiskPoints ($PreviousReport.Exposures | Where-Object { $_.Severity -eq $severity }) } else { $null }
                    Get-TrendHtml -Current $current -Previous $previous -Text "$current pt(s)"
                    New-HTMLText -Alignment center { Get-BadgeHtml -Text $severity -Color $Colors.$severity }
                }
            }
        }

        New-HTMLSection -HeaderText 'Risk points per MITRE ATT&CK tactic' -HeaderBackgroundColor White -HeaderTextColor $Colors.Neutral -CanCollapse -Collapsed {
            $perTactic = $CurrentReport.Exposures | ForEach-Object {
                $exposure = $_
                $_.MITRE -split ', ' | Where-Object { $_ } | ForEach-Object {
                    [PSCustomObject]@{ Tactic = $_; Points = $exposure.Points }
                }
            } | Group-Object Tactic | ForEach-Object {
                [PSCustomObject]@{
                    Tactic = $_.Name
                    Points = [int]($_.Group | Measure-Object -Property Points -Sum).Sum
                    Count  = $_.Count
                }
            } | Where-Object { $_.Points -ne 0 } | Sort-Object -Property Points -Descending
            New-HTMLPanel {
                New-HTMLTable -Title 'Risk points per MITRE ATT&CK tactic' -DataTable $perTactic -PagingLength 10 -HideFooter -HideButtons -DisableSearch
            }
            New-HTMLPanel {
                New-HTMLChart -Title 'Risk points per MITRE ATT&CK tactic' {
                    $i = 0
                    New-ChartLegend -Name $perTactic.Tactic -LegendPosition bottom
                    $perTactic | ForEach-Object {
                        if ($i -ge $Palette.Count) { $i = 0 }
                        New-ChartPie -Value $_.Points -Name $_.Tactic -Color $Palette[$i]
                        $i++
                    }
                }
            }
        }
    }

    # Purple Knight scores per category (0 to 100)
    New-HTMLSection -HeaderText 'Scores per category' {
        $CurrentReport.Categories | Where-Object { $_.Included } | ForEach-Object {
            $category = $_
            New-HTMLPanel {
                New-HTMLGage -Label $category.Name -Value $category.Score
                New-HTMLText -Alignment center -Text "Grade $($category.Grade) - weight $($category.Weight)"
            }
        }
    }

    # Evolution per category between initial, previous and current report
    New-HTMLSection -HeaderText 'Comparison with previous reports' -Direction column {
        New-HTMLText -Text 'The evolution of the Purple Knight score (0 to 100, higher is better) in each category. Ignored indicators (exceptions.csv) do not change these official scores, only the risk points and the exposure tables.'
        New-HTMLSection -Invisible {
            $categoryNames | ForEach-Object {
                $name = $_
                New-HTMLPanel {
                    New-HTMLChart -Title $name {
                        New-ChartBarOptions -Vertical
                        if ($InitialReport) { New-ChartBar -Name 'Initial' -Value ([int]($InitialReport.Categories | Where-Object { $_.Name -eq $name }).Score) }
                        if ($PreviousReport) { New-ChartBar -Name 'Previous' -Value ([int]($PreviousReport.Categories | Where-Object { $_.Name -eq $name }).Score) }
                        New-ChartBar -Name 'Current' -Value ([int]($CurrentReport.Categories | Where-Object { $_.Name -eq $name }).Score)
                    }
                }
            }
        }
    }

    # Comparison between previous and current report
    if ($PreviousReport) {
        New-HTMLSection -HeaderText 'Improvement & deterioration' {
            New-HTMLPanel {
                # The following indicators are not exposed anymore since the last report (improvement)
                New-HTMLSection -Invisible -Margin 0 -AlignItems center -JustifyContent flex-start -BackgroundColor $Colors.Positive {
                    New-HTMLHeading h2 -HeadingText 'Exposures resolved'
                }
                New-HTMLTable -DataTable $riskSolved -DefaultSortIndex 0 -DefaultSortOrder Descending -HideButtons -DisablePaging
            }
            New-HTMLPanel {
                # The following indicators are exposed since the last report (deterioration)
                New-HTMLSection -Invisible -Margin 0 -AlignItems center -JustifyContent flex-start -BackgroundColor $Colors.Negative {
                    New-HTMLHeading h2 -HeadingText 'New exposures'
                }
                New-HTMLTable -DataTable $riskNew -DefaultSortIndex 0 -DefaultSortOrder Descending -HideButtons -DisablePaging
            }
        }
    }

    # Show all exposures
    New-HTMLSection -HeaderText 'All current exposures (indicators of exposure)' -Direction column {
        New-HTMLTable -DataTable (Select-ExposureColumns $CurrentReport.Exposures) -DefaultSortIndex 0 -DefaultSortOrder Descending -DisablePaging
    }

    # Show indicators that could not be evaluated
    if ($CurrentReport.FailedToRun) {
        New-HTMLSection -HeaderText 'Indicators failed to run' -Direction column -CanCollapse {
            New-HTMLText -Text 'These indicators could not be evaluated, their result is unknown.'
            New-HTMLTable -DataTable ($CurrentReport.FailedToRun | Select-Object Severity, ANSSI, Category, IndicatorId, Name, Message) -DefaultSortIndex 0 -DisablePaging
        }
    }

    # Show ignored exposures
    if ($CurrentReport.IgnoredExposures) {
        New-HTMLSection -HeaderText 'Ignored exposures' -Direction column {
            New-HTMLText -Text 'The following indicators have been excluded from the risk points and the ANSSI maturity level using the "exceptions.csv" file.'
            New-HTMLTable -DataTable (Select-ExposureColumns $CurrentReport.IgnoredExposures) -DefaultSortIndex 0 -DefaultSortOrder Descending -DisablePaging
        }
    }
}

function Set-DashboardWidth {
    param (
        [string]$Path,
        [int]$MaxWidth
    )

    $newInlineCss = '<div data-panes="true" style="max-width: ' + $MaxWidth + 'px; margin: 0 auto;">'
    $content = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8) -replace '<div data-panes="true">', $newInlineCss
    # PSWriteHTML serializes missing chart values as "", ApexCharts needs null to leave a gap
    $content = $content -replace '(?<="data":\[[^\]]*)""', 'null'
    [System.IO.File]::WriteAllText($Path, $content, (New-Object System.Text.UTF8Encoding $true))
}

function Get-LogoHtml {
    param (
        [string]$Logo,
        [string]$URI
    )

    if ($Logo) {
        '<a href="{0}" target="_blank"><img src="{1}" style="height: 3em; padding: 1em;" /></a>' -f $URI, $Logo
    }
    else {
        '<span style="color: #fff; font-size: 1.5em; font-weight: bold; padding: 1em;">Purple Knight dashboard</span>'
    }
}
