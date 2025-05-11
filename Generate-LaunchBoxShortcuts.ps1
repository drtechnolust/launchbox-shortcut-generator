# LaunchBox Shortcut Generator
$rootGameFolder = "D:\Arcade\System roms\PC Games 2"
$shortcutOutputFolder = "$rootGameFolder\Shortcuts2"
$maxDepth = 5  # Increased from 3 to 5 to handle deeper folder structures
$folderTimeout = 60  # Maximum seconds to spend on a single game folder
$logFound = "$shortcutOutputFolder\FoundGames.log"
$logNotFound = "$shortcutOutputFolder\NotFoundGames.log"
$logSkipped = "$shortcutOutputFolder\SkippedGames.log"  
$logErrors = "$shortcutOutputFolder\ErrorGames.log"  # New log for error games

# Ensure shortcut output folder exists
if (!(Test-Path $shortcutOutputFolder)) {
    New-Item -ItemType Directory -Path $shortcutOutputFolder | Out-Null
}

# Clear logs if they exist
Remove-Item -Path $logFound, $logNotFound, $logSkipped, $logErrors -ErrorAction SilentlyContinue

# Store a hashtable of created shortcuts to check for duplicates
$createdShortcuts = @{}

function Get-ExecutableScore {
    param($exePath, $gameFolderName)
    $exeName = [System.IO.Path]::GetFileNameWithoutExtension($exePath).ToLower()
    $folderName = $gameFolderName.ToLower()
    $score = 0
    $badNames = @("uninstall", "unins000", "setup", "crashreport", "errorreport", "redist", "redistributable", "vcredist", "directx")
    if ($badNames -contains $exeName) {
        return -1
    }
    
    # Exact match is best
    if ($exeName -eq $folderName) { $score += 10 }
    
    # Partial match is good
    if ($exeName -like "*$folderName*") { $score += 5 }
    
    # Game name as part of the executable name is promising
    $gameNameParts = $folderName -split ' '
    foreach ($part in $gameNameParts) {
        if ($part.Length -gt 3 -and $exeName -like "*$part*") {
            $score += 2
        }
    }
    
    # Priority executable names
    $priorityNames = @("game", "launcher", "start", "play", "run", "main", "bin")
    if ($priorityNames -contains $exeName) { $score += 3 }
    
    # Priority for executables in typical game folders
    $exeLocation = [System.IO.Path]::GetDirectoryName($exePath).ToLower()
    $goodPaths = @("\bin", "\binaries", "\game", "\app", "\win64", "\win32", "\windows", "\x64", "\x86")
    foreach ($goodPath in $goodPaths) {
        if ($exeLocation -like "*$goodPath*") {
            $score += 2
            break
        }
    }
    
    # If it's extremely deep in subfolders, slightly lower priority
    $folderDepth = ($exePath.Split('\').Count - $rootGameFolder.Split('\').Count)
    if ($folderDepth -gt 4) {
        $score -= 1
    }
    
    return $score
}

function Create-Shortcut {
    param (
        [string]$targetExe,
        [string]$shortcutName,
        [string]$outputFolder
    )
    
    # Sanitize the shortcut name by removing/replacing problematic characters
    $safeShortcutName = $shortcutName -replace '[\\\/\:\*\?"<>\|]', '_'  # Replace illegal characters
    $safeShortcutName = $safeShortcutName -replace '&', 'and'  # Replace & with 'and'
    
    # Limit filename length to prevent path too long errors
    if ($safeShortcutName.Length -gt 50) {
        $safeShortcutName = $safeShortcutName.Substring(0, 47) + "..."
    }
    
    # Check if this shortcut already exists based on target executable
    $shortcutPath = "$outputFolder\$safeShortcutName.lnk"
    
    # Check if the shortcut exists and points to the same target
    if (Test-Path $shortcutPath) {
        try {
            $WScriptShell = New-Object -ComObject WScript.Shell
            $existingShortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $existingTarget = $existingShortcut.TargetPath
            
            if ($existingTarget -eq $targetExe) {
                return "exists_same"
            } else {
                return "exists_different"
            }
        } catch {
            Write-Host "Error checking existing shortcut: $_" -ForegroundColor Yellow
            return "error"
        }
    }
    
    # Check if we've already created a shortcut to this executable
    if ($createdShortcuts.ContainsKey($targetExe)) {
        return "duplicate"
    }
    
    # Create the shortcut
    try {
        $WScriptShell = New-Object -ComObject WScript.Shell
        $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $targetExe
        $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetExe)
        $shortcut.Save()
        
        # Add to our tracking hashtable with original name for logging but sanitized for reference
        $createdShortcuts[$targetExe] = $shortcutName
        
        return "created"
    } catch {
        # Try a more aggressive filename sanitization and shorter name
        try {
            $ultraSafeShortcutName = "Game_" + ($shortcutName -replace '[^a-zA-Z0-9]', '_')
            if ($ultraSafeShortcutName.Length -gt 30) {
                $ultraSafeShortcutName = $ultraSafeShortcutName.Substring(0, 30)
            }
            
            $shortcutPath = "$outputFolder\$ultraSafeShortcutName.lnk"
            $WScriptShell = New-Object -ComObject WScript.Shell
            $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = $targetExe
            $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetExe)
            $shortcut.Save()
            
            # Add to our tracking hashtable
            $createdShortcuts[$targetExe] = $shortcutName
            
            # Return special status for sanitized name
            return "created_sanitized"
        } catch {
            Write-Host "Error creating shortcut for $shortcutName`: $_" -ForegroundColor Yellow
            return "error"
        }
    }
}

function Format-TimeSpan {
    param (
        [TimeSpan]$TimeSpan
    )
    
    if ($TimeSpan.TotalHours -ge 1) {
        return "{0:h\:mm\:ss}" -f $TimeSpan
    } else {
        return "{0:mm\:ss}" -f $TimeSpan
    }
}

function Find-Executables {
    param (
        [string]$folderPath,
        [int]$maxDepth,
        [string]$gameName,
        [int]$timeout
    )
    
    $timeoutTime = (Get-Date).AddSeconds($timeout)
    Write-Progress -Id 2 -Activity "Finding executables" -Status "Scanning $gameName..."
    
    $exeFiles = @()
    
    try {
        # IMPORTANT: First check specifically for common game folder structures
        $commonGamePaths = @(
            "$folderPath\Game\*.exe",
            "$folderPath\app\*.exe",
            "$folderPath\bin\*.exe",
            "$folderPath\binaries\*.exe",
            "$folderPath\Windows\*.exe",
            "$folderPath\x64\*.exe",
            "$folderPath\Win64\*.exe",
            "$folderPath\executable\*.exe",
            "$folderPath\program\*.exe",
            "$folderPath\launcher\*.exe",
            "$folderPath\main\*.exe"
        )
        
        foreach ($path in $commonGamePaths) {
            if (Test-Path $path) {
                $foundExes = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                if ($foundExes -and $foundExes.Count -gt 0) {
                    Write-Host "Found executables in common path: $path" -ForegroundColor Green
                    $exeFiles += $foundExes
                }
            }
        }
        
        # Also check for deeper common structures (like Unreal Engine games)
        $deeperCommonPaths = @(
            "$folderPath\Engine\Binaries\Win64\*.exe",
            "$folderPath\Binaries\Win64\*.exe",
            "$folderPath\*\Binaries\Win64\*.exe",
            "$folderPath\*\*\Binaries\Win64\*.exe"
        )
        
        foreach ($path in $deeperCommonPaths) {
            if ($exeFiles.Count -eq 0) {
                try {
                    $foundExes = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                    if ($foundExes -and $foundExes.Count -gt 0) {
                        Write-Host "Found executables in deeper path: $path" -ForegroundColor Green
                        $exeFiles += $foundExes
                    }
                } catch {
                    # Just continue if path is invalid
                }
            }
        }
        
        # If we already found executables in common paths, don't do the expensive searches
        if ($exeFiles.Count -gt 0) {
            Write-Progress -Id 2 -Activity "Finding executables" -Status "Found executables in common paths" -Completed
            return $exeFiles
        }
        
        # Continue with the regular search strategy
        # Try a simple search at the root level first
        $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -ErrorAction SilentlyContinue
        
        # If no files found, search one level deeper
        if ($exeFiles.Count -eq 0) {
            $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -Depth 1 -ErrorAction SilentlyContinue
        }
        
        # If still no files, try a deeper but limited search
        if ($exeFiles.Count -eq 0) {
            $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -Depth $maxDepth -ErrorAction SilentlyContinue
        }
        
        # Final full recursive search if needed, with timeout protection
        if ($exeFiles.Count -eq 0) {
            $scriptBlock = {
                param($path)
                Get-ChildItem -Path $path -Filter "*.exe" -File -Recurse -ErrorAction SilentlyContinue
            }
            
            $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $folderPath
            
            # Wait for job to complete or timeout
            $null = Wait-Job -Job $job -Timeout ($timeout - 5)
            
            if ($job.State -eq "Running") {
                Stop-Job -Job $job
                Write-Host "Timeout reached while scanning $gameName recursively." -ForegroundColor Yellow
            } else {
                $exeFiles = Receive-Job -Job $job
            }
            
            Remove-Job -Job $job -Force
        }
    }
    catch {
        Write-Host "Error scanning $folderPath for executables: $_" -ForegroundColor Red
    }
    
    Write-Progress -Id 2 -Activity "Finding executables" -Completed
    return $exeFiles
}

# Get list of game folders to process
$gameFolders = Get-ChildItem -Path $rootGameFolder -Directory
$total = $gameFolders.Count
$index = 0
$created = 0
$skipped = 0
$notFound = 0
$sanitized = 0
$errors = 0
$timeouts = 0

# Start timing
$startTime = Get-Date
$lastUpdateTime = $startTime

Write-Host "Starting to process $total game folders..."

foreach ($folder in $gameFolders) {
    $index++
    $gameName = $folder.Name
    
    # Calculate timing statistics
    $currentTime = Get-Date
    $elapsedTime = $currentTime - $startTime
    $itemsRemaining = $total - $index
    
    # Only recalculate estimated time every 5 items or every 10 seconds to avoid fluctuations
    if (($index % 5 -eq 0) -or (($currentTime - $lastUpdateTime).TotalSeconds -ge 10)) {
        $lastUpdateTime = $currentTime
        
        if ($index -gt 1) {  # Need at least 2 items to calculate average time
            $averageTimePerItem = $elapsedTime.TotalSeconds / ($index - 1)
            $estimatedTimeRemaining = [TimeSpan]::FromSeconds($averageTimePerItem * $itemsRemaining)
            $formattedTimeRemaining = Format-TimeSpan -TimeSpan $estimatedTimeRemaining
            $formattedElapsedTime = Format-TimeSpan -TimeSpan $elapsedTime
            
            $statusMessage = "$gameName ($index of $total) - $created created, $sanitized sanitized, $skipped skipped, $notFound not found, $errors errors"
            $progressStatus = "$statusMessage | Elapsed: $formattedElapsedTime | Remaining: $formattedTimeRemaining"
        } else {
            $progressStatus = "$gameName ($index of $total)"
        }
    } else {
        $progressStatus = "$gameName ($index of $total) - $created created, $sanitized sanitized, $skipped skipped, $notFound not found, $errors errors"
    }
    
    Write-Progress -Id 1 -Activity "Scanning Games" -Status $progressStatus -PercentComplete (($index / $total) * 100)
    
    # Special handling for known problematic games
    $manualExePath = $null
    if ($gameName -eq "ELDEN RING") {
        $specificPath = "$($folder.FullName)\Game\eldenring.exe"
        if (Test-Path $specificPath) {
            Write-Host "Found Elden Ring executable via direct path!" -ForegroundColor Green
            $manualExePath = $specificPath
        }
    }
    
    # Use our new function to find executables with improved recursive search
    try {
        $exeFiles = @()
        
        # Use manual path if we have one
        if ($manualExePath) {
            $exeFiles = @(Get-Item -Path $manualExePath)
        } else {
            $searchStartTime = Get-Date
            $exeFiles = Find-Executables -folderPath $folder.FullName -maxDepth $maxDepth -gameName $gameName -timeout $folderTimeout
            $searchTime = (Get-Date) - $searchStartTime
            
            # Check if we likely hit a timeout (over 90% of the timeout time used)
            if ($searchTime.TotalSeconds -gt ($folderTimeout * 0.9)) {
                Write-Host "Warning: $gameName search took $($searchTime.TotalSeconds) seconds" -ForegroundColor Yellow
                $timeouts++
                # Log the timeout to errors log
                "$gameName - Timeout reached while scanning for executables" | Out-File -FilePath $logErrors -Append
            }
        }
        
        if ($exeFiles.Count -eq 0) {
            "$gameName - No executable found" | Out-File -FilePath $logNotFound -Append
            $notFound++
            continue
        }
        
        $scored = $exeFiles | ForEach-Object {
            [PSCustomObject]@{
                Path  = $_.FullName
                Score = Get-ExecutableScore -exePath $_.FullName -gameFolderName $gameName
            }
        } | Where-Object { $_.Score -ge 0 } | Sort-Object Score -Descending
        
        if ($scored.Count -eq 0) {
            "$gameName - No suitable exe" | Out-File -FilePath $logNotFound -Append
            $notFound++
            continue
        }
        
        $chosenExe = $scored[0].Path
        $shortcutStatus = Create-Shortcut -targetExe $chosenExe -shortcutName $gameName -outputFolder $shortcutOutputFolder
        
        # Log based on the result
        switch ($shortcutStatus) {
            "created" {
                "$gameName - $chosenExe" | Out-File -FilePath $logFound -Append
                $created++
            }
            "created_sanitized" {
                "$gameName - $chosenExe (with sanitized name)" | Out-File -FilePath $logFound -Append
                $sanitized++
            }
            "exists_same" {
                "$gameName - Shortcut exists and points to same executable: $chosenExe" | Out-File -FilePath $logSkipped -Append
                $skipped++
            }
            "exists_different" {
                "$gameName - Shortcut exists but points to different executable. New: $chosenExe" | Out-File -FilePath $logSkipped -Append
                $skipped++
            }
            "duplicate" {
                "$gameName - Duplicate executable already used for: $($createdShortcuts[$chosenExe])" | Out-File -FilePath $logSkipped -Append
                $skipped++
            }
            "error" {
                "$gameName - Error creating shortcut for: $chosenExe" | Out-File -FilePath $logErrors -Append
                $errors++
            }
        }
    } catch {
        # Log any errors and continue with the next folder
        Write-Host "Error processing $gameName`: $_" -ForegroundColor Red
        "$gameName - Error: $_" | Out-File -FilePath $logErrors -Append
        $errors++
    }
}

$totalTime = (Get-Date) - $startTime
$formattedTotalTime = Format-TimeSpan -TimeSpan $totalTime

Write-Host "✅ Done! Shortcuts saved to: $shortcutOutputFolder"
Write-Host "⏱️ Total time: $formattedTotalTime"
Write-Host "📊 Results: $created created, $sanitized sanitized, $skipped skipped, $notFound not found, $errors errors, $timeouts timeouts"
Write-Host "📄 See FoundGames.log, NotFoundGames.log, SkippedGames.log, and ErrorGames.log for details"