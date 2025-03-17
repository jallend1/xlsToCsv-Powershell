# Create a new Excel instance
$excelApp = New-Object -ComObject Excel.Application

# Turn off Excel alerts
$excelApp.DisplayAlerts = $false

# Iterate through the XLS files in the folder
Get-ChildItem -Path "C:\xlsFiles" -Filter "*.xls" | ForEach-Object {
    # Opens the workbook
    $workbook = $excelApp.Workbooks.Open($_.FullName)

    # Defines the path for the CSV
    $csvFilePath = $_.DirectoryName + "\" + ($_.BaseName + ".csv")
    
    # Save the XLS file as CSV
    $workbook.SaveAs($csvFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV)

    # Close the workbook
    $workbook.Close()

    # Deletes the original XLS file cuz who needs it
    Remove-Item $_.FullName  
}

# Close Excel
$excelApp.Quit()
