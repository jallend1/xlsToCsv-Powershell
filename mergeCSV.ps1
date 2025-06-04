$csvFolder = "C:\csvFiles"
$outputFile = "C:\csvFiles\merged.csv"

# Get all CSV files in the folder
$csvFiles = Get-ChildItem -Path $csvFolder -Filter *.csv

# Get header from the first file, remove the first column
$header = (Get-Content $csvFiles[0].FullName | Select-Object -First 1) -replace '^[^,]+,', ''
$header | Set-Content $outputFile

# Add all rows except the header from each file, remove lines that start with commas, and remove the first column
foreach ($file in $csvFiles) {
    Get-Content $file.FullName |
        Select-Object -Skip 1 |
        Where-Object { -not ($_ -like ",,,,,,,*") } |
        ForEach-Object { $_ -replace '^[^,]+,', '' } |
        Add-Content $outputFile
}

Write-Host "Merged $($csvFiles.Count) files into $outputFile (tracking number column removed)."