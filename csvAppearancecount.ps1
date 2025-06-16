# Merges multiple CSV files and counts unique addresses

$csvFolder = "C:\csvFiles"
$outputFile = "C:\csvFiles\merged_unique.csv"

# Get all CSV files in the folder
$csvFiles = Get-ChildItem -Path $csvFolder -Filter *.csv

# Read all data rows from all files, skipping headers and empty lines
$dataRows = @()
foreach ($file in $csvFiles) {
    $lines = Get-Content $file.FullName | Select-Object -Skip 1
    foreach ($line in $lines) {
        if ($line -and -not ($line -like ",,,,,,,*")) {
            $dataRows += $line
        }
    }
}

# Parse rows into objects (handle quoted fields)
$parsedRows = foreach ($row in $dataRows) {
    $fields = [System.Text.RegularExpressions.Regex]::Split($row, ',(?=(?:[^"]*"[^"]*")*[^"]*$)')
    [PSCustomObject]@{
        'Transaction Date'    = $fields[1]
        'Recipient Name'      = $fields[2]
        'Recipient Company'   = $fields[3]
        'Recipient Address'   = $fields[4]
        'Recipient Country'   = $fields[5]
        'Class'               = $fields[6]
        'Total Cost'          = $fields[7]
        'Package Weight'      = $fields[8]
    }
}

# Group by Recipient Address
$grouped = $parsedRows | Group-Object 'Recipient Address'

# Prepare output with new column, omitting Package Tracking Number
$uniqueRows = foreach ($group in $grouped) {
    $first = $group.Group | Select-Object -First 1
    [PSCustomObject]@{
        'Transaction Date'    = $first.'Transaction Date'
        'Recipient Name'      = $first.'Recipient Name'
        'Recipient Company'   = $first.'Recipient Company'
        'Recipient Address'   = $first.'Recipient Address'
        'Recipient Country'   = $first.'Recipient Country'
        'Class'               = $first.'Class'
        'Total Cost'          = $first.'Total Cost'
        'Package Weight'      = $first.'Package Weight'
        'Appearances'         = $group.Count
    }
}

$uniqueRows | Sort-Object 'Appearances' -Descending | Export-Csv $outputFile -NoTypeInformation

Write-Host "Merged $($csvFiles.Count) files into $outputFile."