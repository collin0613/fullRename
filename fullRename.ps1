<#
fullRename.ps1
Creation Date: 08.20.2024 by Collin Smith
Last Modified: 10.17.2024 by Collin Smith

Purpose: 
    * Combination of rename.ps1 and traverseAndCopy.ps1 *
    Parses through a tree of directories, with each leaf containing a unique set of hearing recording files and a CSV of their case and track numbers.
    First, files are copied from the recorder hard drive and moved to a destination folder on the computer. Files will be deleted from the hard drive.
    Then, checks each target folder to ensure that it meets all the CSV requirements; if not, this target folder will be skipped while the process continues.
    Next, Each file is copied into a temporary folder before being renamed. Renames each file to its corresponding case number by sorting the original names 
    sequentially with the track numbers. Finally, all files are moved into a final Renamed_Files folder, a child of the rootDirPath param folder. If a file 
    has its original name in the Renamed_Files folder, it means that a case number corresponded for two track numbers (two files) and must be manually checked.
    

Usage:
    * ./fullRename.ps1 -source <SOURCE FOLDER PATH> -destination <DESTINATION FOLDER PATH (beyond 2024 parent folder)>
  EX: ./fullRename.ps1 -source "D:/Recordings" -destination "8.20/DM"
    
#>

    param (
        [string]$source,
        [string]$destination
    )

    $startTime = Get-Date
    
    # <<< TODO: CHANGE TO THE PATH WHERE YOU WANT THE LOG FILE TO RESIDE >>>
    $logPath = "C:\Users\colli\OneDrive\Desktop\RenamingLog.txt"

    # Define the names of the leaf folders we are looking for files in
    $targetFolders = @("DM", "EC", "MK", "AR")

    $skippedFolders = New-Object System.Collections.ArrayList

    $dateAndTime = Get-Date -UFormat "%m-%d-%Y %H:%M:%S"

    function Temp-To-Renamed {
        param (
            [string]$tempPath,
            [string]$renamedPath
        )
        $filesProcessed = 0

        # Only used for print messages to the log
        $tempPathShortened = Split-Path $tempPath -Leaf
        $renamedPathShortened = Split-Path $renamedPath -Leaf

        # Check if the paths exist
        if (-not (Test-Path -Path $tempPath)) {
            Write-Error "Temp path does not exist: $tempPath.`nExiting." 2>> $logPath
            exit
        }
        if (-not (Test-Path -Path $renamedPath)) {
            Write-Error "Renamed path does not exist: $renamedPath.`nExiting." 2>> $logPath
            exit
        }

        # Get list of mp3 files in Temp folder
        $files = Get-ChildItem -Path $tempPath -Filter "*.mp3" -File -ErrorAction Stop

        foreach ($file in $files) {
            $originalFileName = $file.Name
            $destinationPath = Join-Path -Path $renamedPath -ChildPath $originalFileName

            $renamedFileList = Get-ChildItem -Path $renamedPath -Filter "*.mp3" -File -ErrorAction Stop

            foreach ($renamedFile in $renamedFileList) {
                if (($renamedFile.Name -eq $file.Name) -and ($renamedFile.Length -eq $file.Length)) {
                    Write-Output "Duplicate file $($file.Name) found in folder $renamedPathShortened. Deleting duplicate from folder $tempPath." | Out-File -FilePath $logPath -Append
                    Remove-Item -Path $file.FullName 
                }
            }

            # Move the file to Renamed_Files
            if (Test-Path -Path $file.FullName) {
                Move-Item -Path $file.FullName -Destination $destinationPath -ErrorAction SilentlyContinue
                $filesProcessed++
                Write-Output "Moved file: $($file.Name) from $tempPathShortened to $renamedPathShortened." | Out-File -FilePath $logPath -Append
            }
        }

        # Check if Temp folder has no more mp3 files
        if (-not (Get-ChildItem -Path $tempPath -Filter "*.mp3" -File)) {
            Remove-Item -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Output "$filesProcessed files were renamed and moved into $renamedPath."
            Write-Output "`nAll available files have been moved/deleted accordingly.`n$filesProcessed files were renamed and moved into $renamedPath.`nDeleted folder: $tempPathShortened." | Out-File -FilePath $logPath -Append
        }
    }


    function Copy-And-Rename {
        param (
        [string]$oldDir,
        [string]$newDir,
        [string]$csvPath
        )
        # Shortened paths to the leaf folder. ONLY used for prints to the log
        $oldDirShortened = Split-Path $oldDir -Leaf
        $newDirShortened = Split-Path $newDir -Leaf

        # Check if the CSV file exists
        if (-Not (Test-Path -Path $csvPath)) {
            Write-Error "The CSV file at the path $csvPath does not exist.`nExiting." 2>> $logPath
            exit
        }

        # Check if source (Original_Files) directory exists
        if (-Not (Test-Path -Path $oldDir)) {
            Write-Error "The source directory $oldDir does not exist.`nExiting." 2>> $logPath
            exit
        }

        # Check if the destination (Renamed_Files) directory exists
        if (-Not (Test-Path -Path $newDir)) {
            try {
                New-Item -Path $newDir -ItemType Directory | Out-Null 
                Write-Output "Created new directory: $newDirShortened" | Out-File -FilePath $logPath -Append
            } catch {
                Write-Error "Failed to create the destination directory $newDir.`nExiting." 2>> $logPath
                exit
            }
        } 

        try {
            # Import the CSV file and sort by trackNumber, numerically ascending
            $csvData = Import-Csv -Path $csvPath
            $sortedData = $csvData | Sort-Object -Property { [int]$_.trackNumber }

        } catch {
            Write-Error "An error occurred while processing the CSV file $csvPath.`nExiting." 2>> $logPath
        }

        # Create the temp folder
        $tempFolderPath = Join-Path -Path $rootDirPath -ChildPath "Temp"
        if (-not (Test-Path -Path $tempFolderPath)) {
            New-Item -ItemType Directory -Path $tempFolderPath
        }
        $tempFolderPathShortened = Split-Path $tempFolderPath -Leaf

        try {
            # Copy all files from source to temp folder
            Copy-Item -Path "$oldDir\*" -Destination $tempFolderPath -Recurse -Force
            Write-Output "All files have been duplicated from $oldDirShortened to $tempFolderPathShortened." | Out-File -FilePath $logPath -Append
        } catch {
            Write-Error "An error occurred while copying files.`nExiting." 2>> $logPath
            exit
        }

        # Get the list of new files in the temp folder
        $files = Get-ChildItem -Path $tempFolderPath -File

        # Gets the list of original files
        $filesOld = Get-ChildItem -Path $dir.FullName -Filter "*.mp3" -File -ErrorAction Stop

        # Check if the number of original files matches the number of rows in the CSV
        if ($filesOld.Count -ne $csvData.Count) {
            Write-Output "The number of original files ($($filesOld.Count)) does not match the number of rows in the CSV ($($csvData.Count)).`nDue to an invalid CSV, this folder will be skipped in the renaming process." | Out-File -FilePath $logPath -Append
            $skippedFolders.Add($dir.FullName)
            continue 
        }

        # Iterate through the files and the new names
        for ($i = 0; $i -lt $filesOld.Count; $i++) {
            $file = $filesOld[$i]
                    
            $newName = $sortedData[$i].caseNumber + ".mp3"
            $tempFolderOldNamePath = Join-Path -Path $tempFolderPath -ChildPath $file.Name

            try {
                # Rename the file
                Rename-Item -Path $tempFolderOldNamePath -NewName $newName
                Write-Output "Renamed copy of $($file.Name) to $newName." | Out-File -FilePath $logPath -Append
            } catch {
                Write-Error "An error occurred while renaming copy of $($file.FullName) to $newFilePath.`nExiting." 2>> $logPath
                exit
            }

            if ($($file.Length) -eq $($newName.Length)) {
                Write-Output "Confirmed that $($file.FullName) and $($newName.FullName) are the same size." | Out-File -FilePath $logPath -Append   
            }
                                 
        }

        continue
    }

    # Function to traverse the directory tree
    function Traverse-Directories {
        param (
            [string]$rootDir
        )

        # Get all subdirectories in the current directory
        $subDirs = Get-ChildItem -Path $rootDir -Directory
 
        foreach ($dir in $subDirs) {

            # If the directory is a target folder AND not a skipped folder
            if (($targetFolders -contains $dir.Name) -and (-not ($skippedFolders -contains $dir.FullName))) {
                Write-Output "`nParsing target folder: $($dir.FullName)" | Out-File -FilePath $logPath -Append

                $subfolderOriginalPath = $dir.FullName
                $subfolderRenamedPath = Join-Path -Path $rootDirPath -ChildPath "Renamed_Files"
                $subfolderRenamedPathDateAdder = Split-Path $rootDirPath -Leaf
                $subfolderRenamedPath = $subfolderRenamedPath + "_" + $subfolderRenamedPathDateAdder
                # Create the folder if it doesn't exist
                if (-not (Test-Path -Path $subfolderRenamedPath)) {
                    New-Item -ItemType Directory -Path $subfolderRenamedPath
                }

                try {
                    # Search for the CSV file in the target folder. Error cases handled previously
                    $csvFile = Get-ChildItem -Path $dir.FullName -Filter "*.csv" -File -ErrorAction Stop
                    $csvFilePath = $csvFile.FullName

                    # Call Copy-And-Rename main function after obtaining the correct path info
                    Copy-And-Rename -oldDir $subfolderOriginalPath -newDir $subfolderRenamedPath -csvPath $csvFilePath

                } catch {
                    Write-Error "Error processing folder $($dir.FullName): $_.`nExiting." 2>> $logPath
                    exit
                }
                continue
            }
            else {
                # Recursively traverse the subdirectory
                Traverse-Directories -rootDir $dir.FullName
            }
        }
    }

    # Function to traverse the directory tree, only checking error cases with the CSV files
    function Traverse-Directories-For-CSV {
        param (
            [string]$rootDir
        )
    
        # Check if the root path exists
        if (-Not (Test-Path -Path $rootDir)) {
            Write-Error "The root path $rootDir does not exist.`nExiting." 2>> $logPath
            exit
        }

        # Get all subdirectories in the current directory
        $subDirs = Get-ChildItem -Path $rootDir -Directory
 
        foreach ($dir in $subDirs) {
            # Skip the directory if it's already in $skippedFolders
            if ($skippedFolders -contains $dir.FullName) {
                continue
            }

            # If the directory is a target folder
            if ($targetFolders -contains $dir.Name) {
                Write-Output "Checking for a valid CSV in target folder: $($dir.FullName)" | Out-File -FilePath $logPath -Append

                try {
                    # Search for the CSV file in the target folder
                    $csvFile = Get-ChildItem -Path $dir.FullName -Filter "*.csv" -File -ErrorAction Stop
                    
                    if ($csvFile.Count -eq 0) {
                        Write-Output "No CSV files exist in $dir.`nDue to an invalid number of CSV files, this folder will be skipped in the renaming process." | Out-File -FilePath $logPath -Append
                        $skippedFolders.Add($dir.FullName)
                        continue
                    }

                    if ($csvFile.Count -gt 1) {
                        Write-Output "Multiple CSV files exist in $dir.`nDue to an invalid number of CSV files, this folder will be skipped in the renaming process." | Out-File -FilePath $logPath -Append
                        $skippedFolders.Add($dir.FullName)
                        continue
                    }

                    $csvPath = $csvFile.FullName

                    # Import the CSV file
                    $csvData = Import-Csv -Path $csvPath

                    # Check if the CSV is empty
                    if ($csvData.Count -eq 0) {
                        Write-Output "The CSV file $csvPath is empty.`nDue to an invalid CSV, this folder will be skipped in the renaming process." | Out-File -FilePath $logPath -Append
                        $skippedFolders.Add($dir.FullName)
                        continue
                    }

                    # Gets the list of files
                    $filesOld = Get-ChildItem -Path $dir.FullName -Filter "*.mp3" -File -ErrorAction Stop

                    # Check if the number of original files matches the number of rows in the CSV
                    if ($filesOld.Count -ne $csvData.Count) {
                        Write-Output "The number of files ($($filesOld.Count)) does not match the number of rows in the CSV ($($csvData.Count)).`nDue to an invalid CSV, this folder will be skipped in the renaming process." | Out-File -FilePath $logPath -Append
                        $skippedFolders.Add($dir.FullName)
                        continue 
                    }
                    
                    # Validate that all trackNumbers are integers
                    foreach ($row in $csvData) {
                        if (-not [int]::TryParse($row.trackNumber, [ref]$null)) {
                            Write-Output "Invalid trackNumber '$($row.trackNumber)' in CSV file $csvPath.`nDue to an invalid trackNumber, this folder will be skipped in the renaming process." | Out-File -FilePath $logPath -Append
                            $skippedFolders.Add($dir.FullName)
                            continue
                        }
                    }

                } catch {
                    Write-Error "Error processing folder $($dir.FullName): $_.`nExiting." 2>> $logPath
                    exit
                }
                continue
            }
            else {
                # Recursively traverse the subdirectory
                Traverse-Directories-For-CSV -rootDir $dir.FullName
            }
        }
    }
    
    function traverseAndCopy {
        param (
            [string]$sourcePath,
            [string]$end
        )

        # Get all subdirectories in the current directory
        $subDirs = Get-ChildItem -Path $sourcePath -Directory
 
        foreach ($dir in $subDirs) {
            $files = Get-ChildItem -Path $dir.FullName -File
            foreach ($file in $files) {
                Copy-Item -Path $file.FullName -Destination $end -Force
                Write-Output "The file $($file.Name) was copied into $end." | Out-File -FilePath $logPath -Append
            }
            traverseAndCopy -sourcePath $dir.FullName -end $end
        }
    }

    function confirmCopy {
        param (
        [string] $folder1,
        [string] $folder2
        )

        $files1 = Get-ChildItem -Path $folder1 -File -Recurse
        $files2 = Get-ChildItem -Path $folder2 -File

        foreach ($file1 in $files1) {
            $matchingFile = $files2 | Where-Object { $_.Name -eq $file1.Name -and $_.Length -eq $file1.Length }
    
            if ($matchingFile) {
                Write-Output "File '$($file1.Name)' exists in both folders with the same size. Deleting the original file." | Out-File -FilePath $logPath -Append
                Remove-Item -Path $file1.FullName -Force
            } else {
                Write-Output "Error: A copy of the file $($file1.Name) was NOT found in $($folder2.Name)" | Out-File -FilePath $logPath -Append
            }
        }
    }

    # Beginning of fullDestination path is hard-coded. The $destination param will be the leaf folder (e.g. "8.20/DM")
    $endPath = "C:\Users\dsmith\desktop\ZoomRecordings\2024"

    Write-Output $dateAndTime | Out-File -FilePath $logPath -Append
    Write-Output "Beginning traversal..." | Out-File -FilePath $logPath -Append
    
    $fullDestination = Join-Path -Path $endPath -ChildPath $destination

    if (-not (Test-Path -Path $fullDestination)) {
        New-Item -Path $fullDestination -ItemType Directory
    }
    if (-not $source) {
        Write-Error "Error: A valid source directory path was not provided." 2>> $logPath
        exit
    } else {
        Write-Output "Now traversing through $source and its subdirectories." | Out-File -FilePath $logPath -Append
        # First stage, copying files from $source (recorder drive) to $destination folder
        traverseAndCopy -sourcePath $source -end $fullDestination
    }

    Write-Output "Now confirming all files were copied:" | Out-File -FilePath $logPath -Append
    confirmCopy -folder1 $source -folder2 $fullDestination
    # After confirmCopy, the traverseAndCopy.ps1 portion of the script is done, and the rename.ps1 portion of the script begins

    # rootDirPath is not a param like it was in rename.ps1, but is the parent of the destination param (e.g. "8.20/DM" --> "8.20")
    $rootDirPath = Split-Path -Path $fullDestination -Parent

    # Enters csv checking stage, parsing to check valid CSV files
    Traverse-Directories-For-CSV -rootDir $rootDirPath

    if ($skippedFolders.Count -gt 0) {
        $problemPath = Join-Path -Path $rootDirPath -ChildPath "problem.txt"
        if (-not (Test-Path -Path $problemPath)) {
            New-Item -Path $problemPath -ItemType File | Out-Null
        }
        Write-Output $dateAndTime | Out-File -FilePath $problemPath -Append
        Write-Output "Skipped Folders: $($skippedFolders -join ', ')`n" | Out-File -FilePath $problemPath -Append
        Write-Output "Skipped Folders: $($skippedFolders -join ', ')" | Out-File -FilePath $logPath -Append
    }

    # Enters next stage, parsing to rename files
    Write-Output "`nBeginning Copy-And-Rename process..." | Out-File -FilePath $logPath -Append
    Traverse-Directories -rootDir $rootDirPath

    $renamedPathDateAdder = Split-Path $rootDirPath -Leaf
    $renamedPath = Join-Path -Path $rootDirPath -ChildPath "Renamed_Files"
    $renamedPath = $renamedPath + "_" + $renamedPathDateAdder

    # Enters next stage
    Write-Output "`nFiles have been copied into Temp folder and renamed accordingly.`nEntering final file names into $renamedPath...`n" | Out-File -FilePath $logPath -Append

    if (-not (Test-Path -Path $renamedPath)) {
        New-Item -ItemType Directory -Path $renamedPath
    }
    $temporaryPath = Join-Path -Path $rootDirPath -ChildPath "Temp"
    if (-not (Test-Path -Path $temporaryPath)) {
        New-Item -ItemType Directory -Path $renamedPath
    }

    Temp-To-Renamed -tempPath $temporaryPath -renamedPath $renamedPath

    $endTime = Get-Date
    $duration = $endTime - $startTime

    # Writes a successful confirmation line at the end of the log and to the console
    Write-Output "Script completed in $($duration.TotalSeconds) seconds.`nComplete.`n`n" | Out-File -FilePath $logPath -Append
    Write-Output "Script completed in $($duration.TotalSeconds) seconds."