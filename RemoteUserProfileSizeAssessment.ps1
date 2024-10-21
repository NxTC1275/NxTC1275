# Prompt user for output directory
$outputDir = Read-Host "Please enter the folder path where results will be saved"

# Ensure the directory exists
if (-Not (Test-Path -Path $outputDir)) {
    Write-Host "The specified directory does not exist. Exiting script." -ForegroundColor Red
    exit
}

# Define the files for CSV and HTML output
$csvFile = Join-Path -Path $outputDir -ChildPath "OrphanedContent.csv"
$htmlFile = Join-Path -Path $outputDir -ChildPath "OrphanedContent.html"

# Set a reasonable timeout value
$timeout = 10 # 10 seconds for each connection

# Get the list of computers in the domain
$computers = Get-ADComputer -Filter * | Select-Object -ExpandProperty Name

# Initialize an empty array for results
$results = @()

# Loop through each computer and process user profiles
foreach ($computer in $computers) {
    Write-Verbose "Processing computer: $computer"
    
    $adminShare = "\\$computer\C$"
    $usersPath = Join-Path -Path $adminShare -ChildPath "Users"

    # Retry connection to computer up to 3 times
    $pingSuccess = $false
    for ($i = 0; $i -lt 3; $i++) {
        try {
            $ping = Test-Connection -ComputerName $computer -Count 1 -ErrorAction Stop
            if ($ping.StatusCode -eq 0) {
                $pingSuccess = $true
                break
            }
        } catch {
            Write-Warning "Attempt $($i + 1) to connect to $computer failed."
            Start-Sleep -Seconds 5  # Wait for 5 seconds before retrying
        }
    }

    # If the ping fails after retries, move to the next computer
    if (-not $pingSuccess) {
        Write-Warning "Failed to connect to $computer after 3 attempts."
        continue  # Skip to the next computer
    }

    try {
        # Get all user profiles from the Users folder
        $userFolders = Get-ChildItem -Path $usersPath -ErrorAction Stop | Where-Object { $_.PSIsContainer }

        foreach ($userFolder in $userFolders) {
            Write-Verbose "Processing user folder: $($userFolder.Name)"

            $subfolders = @("Desktop", "Documents", "Downloads", "Favorites", "Pictures", "Videos")

            foreach ($subfolder in $subfolders) {
                $subfolderPath = Join-Path -Path $userFolder.FullName -ChildPath $subfolder

                if (Test-Path $subfolderPath) {
                    $size = (Get-ChildItem -Recurse -File $subfolderPath | Measure-Object -Property Length -Sum).Sum

                    # Collect result for CSV
                    $results += [pscustomobject]@{
                        Computer     = $computer
                        User         = $userFolder.Name
                        Folder       = $subfolder
                        SizeInBytes  = $size
                    }

                    Write-Verbose "$subfolder folder size for $($userFolder.Name): $size bytes"
                }
            }
        }
    } catch {
        Write-Warning "Failed to connect to $computer or fetch user profiles: $_"
    }
}

# Export to CSV
$results | Export-Csv -Path $csvFile -NoTypeInformation

# Convert to HTML for readable format
$resultsHtml = $results | ConvertTo-Html -Property Computer, User, Folder, SizeInBytes -PreContent "<h1>Orphaned Content Report</h1>" -PostContent "<p>Generated on $(Get-Date)</p>"

# Save the HTML report
$resultsHtml | Out-File -FilePath $htmlFile

Write-Host "Report saved to $csvFile and $htmlFile"

