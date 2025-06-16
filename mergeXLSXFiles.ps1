# Combine all XLSX files in a folder, excluding rows with an empty address

# The folder containing the XLSX files
$folder = "C:\xlsFiles"
# Output CSV file path
$outputCsv = "C:\xlsFiles\combined.csv"

# Install ImportExcel module if not present
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Get all XLSX files in the folder
$files = Get-ChildItem -Path $folder -Filter *.xlsx

# Initialize an array to hold all rows
$allRows = @()

foreach ($file in $files) {
    # Import the XLSX file
    $rows = Import-Excel -Path $file.FullName

    # Filter out rows where "Recipient Address" is null or empty
    $filteredRows = $rows | Where-Object { $_."Recipient Address" -ne $null -and $_."Recipient Address".Trim() -ne "" }

    $allRows += $filteredRows
}

# Export the combined rows to CSV
$allRows | Export-Csv -Path $outputCsv -NoTypeInformation

Write-Host "Combined file saved to $outputCsv"