# Combine all XLSX files in a folder, excluding rows with an empty address

# The folder containing the XLSX files
$folder = "C:\updatedPost"
# Output CSV file path
$outputCsv = "C:\updatedPost\combined.csv"

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

    # If "Shipment Date" exists, copy its value to "Transaction Date" for each row
    if ($rows | Get-Member -Name "Shipment Date" -MemberType NoteProperty) {
        foreach ($row in $rows) {
            $transactionDate = $row."Shipment Date"
            # Returns error saying 'Transaction Date' cannot be found, so I'm always adding it
            $row | Add-Member -NotePropertyName "Transaction Date" -NotePropertyValue $transactionDate -Force
            $row.PSObject.Properties.Remove("Shipment Date")
        }
    }

    # Filter out rows where "Recipient Address" is empty to eliminate mysterious summary rows
    $filteredRows = $rows | Where-Object { $_."Recipient Address" -ne $null -and $_."Recipient Address".Trim() -ne "" }

    $allRows += $filteredRows
}

# Export the combined rows to CSV
$allRows | Export-Csv -Path $outputCsv -NoTypeInformation

Write-Host "Combined file saved to $outputCsv"