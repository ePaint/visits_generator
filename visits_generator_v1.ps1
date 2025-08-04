try {
    Import-Module powershell-yaml -ErrorAction Stop
} catch {
    Write-Host "Error: powershell-yaml module is not installed. Please install it using " -ForegroundColor Red -NoNewLine
    Write-Host "Install-Module powershell-yaml" -ForegroundColor Cyan
    exit 1
}

$settings = Get-Content -Path "$PSScriptRoot\settings.yaml" | ConvertFrom-Yaml
$inputFiles = Get-ChildItem -Path $settings.input_folder -Filter "*.csv" | Where-Object { $_.Name -match $settings.filename_format_regex }
if ($inputFiles.Count -eq 0) {
    Write-Host "No files found matching the pattern $($settings.filename_format_regex) in $($settings.input_folder)." -ForegroundColor Yellow
    exit 0
}

$culture = [System.Globalization.CultureInfo]::CreateSpecificCulture('en-US')
$minTime = Get-Date $settings.min_time
$maxTime = Get-Date $settings.max_time
$totalSeconds = ($maxTime - $minTime).TotalSeconds

foreach ($file in $inputFiles) {
    Write-Host "Processing file: $($file.Name)" -ForegroundColor Cyan
    $data = Import-Csv -Path $file.FullName
    $fileDate = $file.Name -replace $settings.filename_format_regex, '$1'
    $fileDate = [datetime]::ParseExact($fileDate, $settings.date_format, [System.Globalization.CultureInfo]::InvariantCulture)
    $validStartDate = (Get-Date -Year $fileDate.Year -Month $fileDate.Month -Day 1)
    $validEndDate = (Get-Date -Year $fileDate.Year -Month $fileDate.Month -Day 1).AddMonths(1).AddDays(-1)
    $totalDays = ($validEndDate - $validStartDate).Days + 1
    Write-Host "Valid date range: $validStartDate to $validEndDate" -ForegroundColor Green

    $entries = @{}
    $data | ForEach-Object {
        $Name = "$($_.'First Name') $($_.'Last Name')"
        if (-not $entries.ContainsKey($Name)) {
            Write-Host "Processing visitor: '$Name'" -ForegroundColor Green
            $entries[$Name] = @{
                'Account Number' = $_.'Account Number'
                'ID Number' = $_.'ID Number'
                'First Name' = $_.'First Name'
                'Last Name' = $_.'Last Name'
                'Program' = $_.'Program'
                'Visits' = @()
            }
            if ($settings.entries_per_visitor.ContainsKey($Name)) {
                $minVisits = $settings.entries_per_visitor[$Name].min
                $maxVisits = $settings.entries_per_visitor[$Name].max
            }
            elseif ($settings.ask_for_missing_entries -eq $true) {
                $minVisits = Read-Host "Enter minimum visits for $Name (default: $($settings.unknown_visitor_min_entries))"
                $maxVisits = Read-Host "Enter maximum visits for $Name (default: $($settings.unknown_visitor_max_entries))"
                if (-not [int]::TryParse($minVisits, [ref]$minVisits)) {
                    $minVisits = $settings.unknown_visitor_min_entries
                }
                if (-not [int]::TryParse($maxVisits, [ref]$maxVisits)) {
                    $maxVisits = $settings.unknown_visitor_max_entries
                }
            } else {
                $minVisits = $settings.unknown_visitor_min_entries
                $maxVisits = $settings.unknown_visitor_max_entries
            }
            $entries[$Name].minVisits = $minVisits
            $entries[$Name].maxVisits = $maxVisits
            
        }
        $entries[$Name].Visits += @{
            'Check-In Date' = $_.'Check-In Date'
            'Check-In Time' = $_.'Check-In Time'
        }
    }
    foreach ($Name in $entries.Keys) {
        if ($entries[$Name].Visits.Count -lt $entries[$Name].minVisits) {
            $numberOfVisits = Get-Random -Minimum $entries[$Name].minVisits -Maximum ($entries[$Name].maxVisits + 1)
            $visitsToAdd = $numberOfVisits - $entries[$Name].Visits.Count
            Write-Host "Warning: $Name has less than the minimum required visits. Current: $($entries[$Name].Visits.Count). Minimum: $($entries[$Name].minVisits). Maximum: $($entries[$Name].maxVisits). Adding: $visitsToAdd. New Total: $numberOfVisits" -ForegroundColor Yellow
            $maxRetries = 1000
            $tries = 0
            for ($i = 0; $i -lt $visitsToAdd; $i++) {
                $randomDate = $validStartDate.AddDays((Get-Random -Minimum 0 -Maximum $totalDays))
                if ($settings.valid_days -and -not $settings.valid_days.Contains($randomDate.DayOfWeek.ToString())) {
                    $i--
                    $tries++
                    if ($tries -ge $maxRetries) {
                        Write-Host "Error: Could not find a unique date for $Name after $maxRetries attempts. Skipping." -ForegroundColor Red
                        exit 1
                    }
                    continue
                }
                if (-not $settings.can_repeat_days -and ($entries[$Name].Visits | Where-Object { $_.'Check-In Date' -eq $randomDate.ToString("yyyy-MM-dd") })) {
                    $i--
                    $tries++
                    if ($tries -ge $maxRetries) {
                        Write-Host "Error: Could not find a unique date for $Name after $maxRetries attempts. Skipping." -ForegroundColor Red
                        exit 1
                    }
                    continue
                }
                $randomTime = $minTime.Add([timespan]::FromSeconds((Get-Random -Minimum 0 -Maximum $totalSeconds)))
                $entries[$Name].Visits += @{
                    'Check-In Date' = $randomDate.ToString("yyyy-MM-dd")
                    'Check-In Time' = $randomTime.ToString("h:mmtt", $culture)
                }
            }
        } elseif ($entries[$Name].Visits.Count -gt $entries[$Name].maxVisits) {
            Write-Host "Warning: $Name has more than the maximum allowed visits. Current: $($entries[$Name].Visits.Count). Maximum: $($entries[$Name].maxVisits)" -ForegroundColor Gray
        }
    }

    $data = @()
    foreach ($entry in $entries.GetEnumerator()) {
        $Name = $entry.Key
        $visitor = $entry.Value
        $entries[$Name].Visits = $entries[$Name].Visits | Sort-Object { $_.'Check-In Date' }, { $_.'Check-In Time' }
        foreach ($visit in $visitor.Visits) {
            $data += [PSCustomObject]@{
                'Account Number' = $visitor.'Account Number'
                'ID Number' = $visitor.'ID Number'
                'First Name' = $visitor.'First Name'
                'Last Name' = $visitor.'Last Name'
                'Program' = $visitor.'Program'
                'Check-In Date' = $visit.'Check-In Date'
                'Check-In Time' = $visit.'Check-In Time'.ToLower()
                'Total Visits' = $visitor.Visits.Count
            }
        }
    }

    $outputFileName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + $settings.output_filename_suffix + ".csv"
    $outputFilePath = Join-Path -Path $settings.output_folder -ChildPath $outputFileName
    $data | Export-Csv -Path $outputFilePath -NoTypeInformation -Encoding UTF8 -Force

    Write-Host "Processed file saved to: $outputFilePath" -ForegroundColor Cyan
}
