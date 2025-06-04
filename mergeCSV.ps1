$csvFolder = "C:\csvFiles"
$outputFile = "C:\csvFiles\merged.csv"

# Get all CSV files in the folder
$csvFiles = Get-ChildItem -Path $csvFolder -Filter *.csv

# Read the header from the first file
$header = Get-Content $csvFiles[0].FullName | Select-Object -First 1

# Write the header to the output file
$header | Set-Content $outputFile

# Append all rows except the header from each file
foreach ($file in $csvFiles) {
    Get-Content $file.FullName | Select-Object -Skip 1 | Add-Content $outputFile
}

Write-Host "Merged $($csvFiles.Count) files into $outputFile"