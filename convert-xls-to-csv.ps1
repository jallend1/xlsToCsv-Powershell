# Converts a folder of XLS files to CSV format

# Create a new Excel instance
$excelApp = New-Object -ComObject Excel.Application

# Turn off Excel alerts
$excelApp.DisplayAlerts = $false

# Define source and destination folders
$sourceFolder = "C:\xlsFiles"
$destFolder = "C:\csvFiles"

# Create destination folder if it doesn't exist
if (!(Test-Path -Path $destFolder)) {
    New-Item -ItemType Directory -Path $destFolder | Out-Null
}

# Iterate through the XLS files in the folder
Get-ChildItem -Path $sourceFolder -Filter "*.xls" | ForEach-Object {
    Write-Host "Processing file: $($_.FullName)"

    # Opens the workbook
    $workbook = $excelApp.Workbooks.Open($_.FullName)

    # Defines the path for the CSV in the destination folder
    $csvFilePath = Join-Path $destFolder ($_.BaseName + ".csv")
    
    # Save the XLS file as CSV
    $workbook.SaveAs($csvFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV)
    
    Write-Host "Converted to CSV: $csvFilePath"

    # Close the workbook
    $workbook.Close()
    
    # Deletes the original XLS file if desired
    # Remove-Item $_.FullName  
}

Write-Host "All files processed."

# Close Excel
$excelApp.Quit()