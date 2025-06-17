# Set source and destination folders
$sourceFolder = "C:\tsvSource"
$destFolder = "C:\tsvConverted"

# Creates destination folder if it doesn't exist
if (-not (Test-Path $destFolder)) {
    New-Item -ItemType Directory -Path $destFolder | Out-Null
}

# Get all TSV files in the source folder
$tsvFiles = Get-ChildItem -Path $sourceFolder -Filter *.tsv

foreach ($file in $tsvFiles) {
    $tsvPath = $file.FullName
    $csvPath = Join-Path $destFolder ([System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".csv")

    # TODO: Establish header mapping
    Import-Csv -Path $tsvPath -Delimiter "`t" | Export-Csv -Path $csvPath -NoTypeInformation
}

Write-Host "All TSV files have been converted to CSV in $destFolder."