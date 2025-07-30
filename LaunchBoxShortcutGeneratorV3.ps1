# LaunchBox Shortcut Generator V3 - Enhanced Version with User Input
# This script scans a games folder and creates Windows shortcuts for each game
# by intelligently identifying the main executable file in each game's folder

# ====================================================================================
# USER INPUT SECTION - Get and validate the root game folder path
# ====================================================================================

# Get user input for the root game folder with validation loop
do {
    $rootGameFolder = Read-Host "Enter the path to your games folder"
    
    # Validate that user entered something
    if (-not $rootGameFolder) {
        Write-Host "Please enter a valid path." -ForegroundColor Red
        continue
    }
    
    # Remove quotes if user included them (common when copy-pasting paths)
    $rootGameFolder = $rootGameFolder.Trim('"')
    
    # Check if the path actually exists on the filesystem
    if (-not (Test-Path $rootGameFolder)) {
        Write-Host "Path does not exist: $rootGameFolder" -ForegroundColor Red
        Write-Host "Please enter a valid folder path." -ForegroundColor Yellow
        continue
    }
    
    # Check if it's actually a directory (not a file)
    if (-not (Get-Item $rootGameFolder).PSIsContainer) {
        Write-Host "Path is not a directory: $rootGameFolder" -ForegroundColor Red
        continue
    }
    
    # If we get here, the path is valid - break out of the loop
    break
} while ($true)

Write-Host "Scanning folder: $rootGameFolder" -ForegroundColor Green

# ====================================================================================
# CONFIGURATION VARIABLES - Script behavior settings
# ====================================================================================

# Where to save the generated shortcuts
$shortcutOutputFolder = "$rootGameFolder\Shortcuts"

# Maximum folder depth to search for executables (prevents infinite recursion)
$maxDepth = 8  # Increased from 5 to 8 to handle deeper folder structures

# Maximum time to spend scanning each game folder (prevents hanging on problematic folders)
$folderTimeout = 300  # Increased from 60 to 300 seconds per game folder

# Log file paths for different types of results
$logFound = "$shortcutOutputFolder\FoundGames.log"        # Successfully created shortcuts
$logNotFound = "$shortcutOutputFolder\NotFoundGames.log"  # No suitable executable found
$logSkipped = "$shortcutOutputFolder\SkippedGames.log"    # Shortcuts already exist or duplicates
$logErrors = "$shortcutOutputFolder\ErrorGames.log"       # Errors during processing
$logBlocked = "$shortcutOutputFolder\BlockedFiles.log"    # Files blocked by filters (for debugging)

# Ensure shortcut output folder exists
if (!(Test-Path $shortcutOutputFolder)) {
    New-Item -ItemType Directory -Path $shortcutOutputFolder | Out-Null
}

# Clear logs if they exist (start fresh each run)
Remove-Item -Path $logFound, $logNotFound, $logSkipped, $logErrors, $logBlocked -ErrorAction SilentlyContinue

# Store a hashtable of created shortcuts to check for duplicates
# Key: executable path, Value: game name
$createdShortcuts = @{}

# ====================================================================================
# EXECUTABLE SCORING FUNCTION - Determines which .exe is most likely the main game
# ====================================================================================

function Get-ExecutableScore {
    param($exePath, $gameFolderName)
    $exeName = [System.IO.Path]::GetFileNameWithoutExtension($exePath).ToLower()
    $folderName = $gameFolderName.ToLower()
    $score = 0
    
    # Phase 1: Block files with these problematic words anywhere in the name
    # These patterns indicate utility/setup files rather than the main game executable
    $blockPatterns = @(
        "*uninstall*", "*setup*", "*settings*", "*helper*",
        "*config*", "*launcher*", "*language*", "*crash*", 
        "*test*", "*service*", "*server*", "*update*", "*install*"
    )
    foreach ($pattern in $blockPatterns) {
        if ($exeName -like $pattern) {
            # Log blocked files for debugging - helps identify false positives
            "$gameFolderName - BLOCKED by pattern '$pattern': $exePath" | Out-File -FilePath $logBlocked -Append
            return -1  # Negative score = excluded from consideration
        }
    }
    
    # Enhanced blacklist with additional exclusions for known non-game executables
    # These are exact filename matches (without extension) that should never be selected
    $badNames = @(
        "unins000",  # Common uninstaller name
        "crashreport", "errorreport", "crashreporter", "crashpad",  # Crash reporting tools
        "redist", "redistributable", "vcredist", "vc_redist",       # Microsoft redistributables
        "directx", "dxwebsetup",                                    # DirectX installers
        "uploader", "webhelper",                                    # Web/upload utilities
        "crs-handler", "crs-uploader", "crs-video",                # CRS (crash reporting) tools
        "drivepool", "quicksfv", "handler",                        # Utility programs
        "gamingrepair", "unitycrashhandle64"                       # Repair/crash handling tools
    )
    
    if ($badNames -contains $exeName) {
        # Log blocked files for debugging
        "$gameFolderName - BLOCKED by blacklist: $exePath" | Out-File -FilePath $logBlocked -Append
        return -1  # Exclude from consideration
    }
    
    # HIGHEST PRIORITY: Exact match with folder name gets maximum score
    # If folder is "Half-Life 2" and exe is "half-life 2.exe", this is almost certainly correct
    if ($exeName -eq $folderName) { 
        return 1000  # Guaranteed highest score
    }
    
    # HIGH PRIORITY: Check for shipping executables (Unreal Engine pattern)
    # Many Unreal Engine games use pattern like "gamename-win64-shipping.exe"
    $shippingPattern = "$folderName-win64-shipping"
    if ($exeName -eq $shippingPattern) {
        return 500  # Very high score, but less than exact match
    }
    
    # Also check for variations of shipping executables with partial matches
    if ($exeName -like "*$folderName*" -and $exeName -like "*win64*shipping*") {
        $score += 100
    }
    
    # ENHANCED: Check for "game" anywhere in executable name
    # Fixes cases like gsgameexe.exe, gameexe.exe, etc. where "game" indicates main executable
    if ($exeName -like "*game*") { 
        $score += 75  # High bonus for any game-related executable
    }
    
    # Partial match with folder name is promising
    # If folder is "Skyrim" and exe is "tesv_skyrim.exe", this gets bonus points
    if ($exeName -like "*$folderName*") { $score += 50 }
    
    # Game name as part of the executable name is promising
    # Break folder name into words and see if any appear in the executable name
    $gameNameParts = $folderName -split ' '
    foreach ($part in $gameNameParts) {
        if ($part.Length -gt 3 -and $exeName -like "*$part*") {
            $score += 20  # Bonus for each significant word match
        }
    }
    
    # Priority executable names that commonly indicate main game files
    # Note: "launcher" was removed from this list since it's blocked above
    $priorityNames = @("start", "play", "run", "main", "bin")
    if ($priorityNames -contains $exeName) { $score += 30 }
    
    # Priority for executables in typical game folders
    # Games often put their main executable in these common subdirectories
    $exeLocation = [System.IO.Path]::GetDirectoryName($exePath).ToLower()
    $goodPaths = @("\bin", "\binaries", "\game", "\app", "\win64", "\win32", "\windows", "\x64", "\x86")
    foreach ($goodPath in $goodPaths) {
        if ($exeLocation -like "*$goodPath*") {
            $score += 20  # Bonus for being in a typical game directory
            break
        }
    }
    
    # If it's extremely deep in subfolders, slightly lower priority
    # Very deeply nested files are often utilities rather than main executables
    $folderDepth = ($exePath.Split('\').Count - $rootGameFolder.Split('\').Count)
    if ($folderDepth -gt 6) {
        $score -= 10  # Small penalty for being too deep
    }
    
    return $score
}

# ====================================================================================
# SHORTCUT CREATION FUNCTION - Creates Windows .lnk files safely
# ====================================================================================

function Create-Shortcut {
    param (
        [string]$targetExe,      # Path to the executable to create shortcut for
        [string]$shortcutName,   # Name for the shortcut (usually the game folder name)
        [string]$outputFolder    # Where to save the .lnk file
    )
    
    # Sanitize the shortcut name by removing/replacing problematic characters
    # Windows filenames cannot contain these characters: \/:*?"<>|
    $safeShortcutName = $shortcutName -replace '[\\\/\:\*\?"<>\|]', '_'  # Replace illegal characters
    $safeShortcutName = $safeShortcutName -replace '&', 'and'  # Replace & with 'and' for better compatibility
    
    # Limit filename length to prevent "path too long" errors
    if ($safeShortcutName.Length -gt 50) {
        $safeShortcutName = $safeShortcutName.Substring(0, 47) + "..."
    }
    
    # Build the full path for the shortcut file
    $shortcutPath = "$outputFolder\$safeShortcutName.lnk"
    
    # Check if this shortcut already exists based on target executable
    if (Test-Path $shortcutPath) {
        try {
            # Check what the existing shortcut points to
            $WScriptShell = New-Object -ComObject WScript.Shell
            $existingShortcut = $WScriptShell.CreateShortcut($shortcutPath)
            $existingTarget = $existingShortcut.TargetPath
            
            # If it points to the same executable, no need to recreate
            if ($existingTarget -eq $targetExe) {
                return "exists_same"
            } else {
                return "exists_different"  # Points to different executable
            }
        } catch {
            Write-Host "Error checking existing shortcut: $_" -ForegroundColor Yellow
            return "error"
        }
    }
    
    # Check if we've already created a shortcut to this executable
    # Prevents creating multiple shortcuts to the same game executable
    if ($createdShortcuts.ContainsKey($targetExe)) {
        return "duplicate"
    }
    
    # Create the shortcut using Windows Script Host COM object
    try {
        $WScriptShell = New-Object -ComObject WScript.Shell
        $shortcut = $WScriptShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $targetExe  # What the shortcut points to
        $shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($targetExe)  # Set working directory
        $shortcut.Save()  # Actually create the .lnk file
        
        # Add to our tracking hashtable with original name for logging
        $createdShortcuts[$targetExe] = $shortcutName
        
        return "created"
    } catch {
        # Try a more aggressive filename sanitization and shorter name if first attempt fails
        try {
            # Ultra-safe naming: only alphanumeric characters and underscores
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

# ====================================================================================
# TIME FORMATTING FUNCTION - Makes elapsed time human-readable
# ====================================================================================

function Format-TimeSpan {
    param (
        [TimeSpan]$TimeSpan
    )
    
    # Format as h:mm:ss if over an hour, otherwise just mm:ss
    if ($TimeSpan.TotalHours -ge 1) {
        return "{0:h\:mm\:ss}" -f $TimeSpan
    } else {
        return "{0:mm\:ss}" -f $TimeSpan
    }
}

# ====================================================================================
# EXECUTABLE SEARCH FUNCTION - Finds .exe files with optimized search strategy
# ====================================================================================

function Find-Executables {
    param (
        [string]$folderPath,     # Game folder to search
        [int]$maxDepth,          # Maximum recursion depth
        [string]$gameName,       # Game name for logging
        [int]$timeout            # Maximum seconds to spend searching
    )
    
    # Set up timeout protection
    $timeoutTime = (Get-Date).AddSeconds($timeout)
    Write-Progress -Id 2 -Activity "Finding executables" -Status "Scanning $gameName..."
    
    $exeFiles = @()
    
    try {
        # IMPORTANT: First check specifically for common game folder structures
        # Many games follow standard patterns, so check these first for efficiency
        $commonGamePaths = @(
            "$folderPath\Game\*.exe",        # Common in many game engines
            "$folderPath\app\*.exe",         # Steam and other platforms
            "$folderPath\bin\*.exe",         # Binary/executable folder
            "$folderPath\binaries\*.exe",    # Alternative binary folder
            "$folderPath\Windows\*.exe",     # Platform-specific folder
            "$folderPath\x64\*.exe",         # 64-bit executables
            "$folderPath\Win64\*.exe",       # 64-bit Windows executables
            "$folderPath\executable\*.exe",  # Some games use this
            "$folderPath\program\*.exe",     # Program files folder
            "$folderPath\launcher\*.exe",    # Launcher folder
            "$folderPath\main\*.exe"         # Main executable folder
        )
        
        # Search each common path pattern
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
        # Unreal Engine games often have deep nested structure
        $deeperCommonPaths = @(
            "$folderPath\Engine\Binaries\Win64\*.exe",    # Unreal Engine structure
            "$folderPath\Binaries\Win64\*.exe",           # Alternative Unreal structure
            "$folderPath\*\Binaries\Win64\*.exe",         # One level deeper
            "$folderPath\*\*\Binaries\Win64\*.exe",       # Two levels deeper
            "$folderPath\*\*\*\Binaries\Win64\*.exe"      # Three levels deeper
        )
        
        # Only check deeper paths if we haven't found anything yet
        foreach ($path in $deeperCommonPaths) {
            if ($exeFiles.Count -eq 0) {
                try {
                    $foundExes = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                    if ($foundExes -and $foundExes.Count -gt 0) {
                        Write-Host "Found executables in deeper path: $path" -ForegroundColor Green
                        $exeFiles += $foundExes
                    }
                } catch {
                    # Just continue if path is invalid - wildcards can cause issues
                }
            }
        }
        
        # If we already found executables in common paths, don't do the expensive searches
        # This is a major performance optimization
        if ($exeFiles.Count -gt 0) {
            Write-Progress -Id 2 -Activity "Finding executables" -Status "Found executables in common paths" -Completed
            return $exeFiles
        }
        
        # Continue with the regular search strategy if common paths didn't work
        # Progressive search: start shallow and go deeper only if needed
        
        # Try a simple search at the root level first
        $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -ErrorAction SilentlyContinue
        
        # If no files found, search one level deeper
        if ($exeFiles.Count -eq 0) {
            $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -Depth 1 -ErrorAction SilentlyContinue
        }
        
        # If still no files, try progressively deeper searches
        if ($exeFiles.Count -eq 0) {
            # Try depth 2
            $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -Depth 2 -ErrorAction SilentlyContinue
        }
        
        if ($exeFiles.Count -eq 0) {
            # Try depth 4
            $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -Depth 4 -ErrorAction SilentlyContinue
        }
        
        # If still no files, try the full max depth search
        if ($exeFiles.Count -eq 0) {
            $exeFiles = Get-ChildItem -Path $folderPath -Filter "*.exe" -File -Depth $maxDepth -ErrorAction SilentlyContinue
        }
        
        # Final full recursive search if needed, with timeout protection
        # This is the most expensive operation, so we do it last and with timeout
        if ($exeFiles.Count -eq 0) {
            # Use PowerShell job to enable timeout capability
            $scriptBlock = {
                param($path)
                Get-ChildItem -Path $path -Filter "*.exe" -File -Recurse -ErrorAction SilentlyContinue
            }
            
            # Start the search as a background job
            $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $folderPath
            
            # Wait for job to complete or timeout (leave 5 seconds buffer)
            $null = Wait-Job -Job $job -Timeout ($timeout - 5)
            
            # Check if job is still running (timed out)
            if ($job.State -eq "Running") {
                Stop-Job -Job $job  # Kill the job
                Write-Host "Timeout reached while scanning $gameName recursively." -ForegroundColor Yellow
            } else {
                # Job completed, get the results
                $exeFiles = Receive-Job -Job $job
            }
            
            # Clean up the job
            Remove-Job -Job $job -Force
        }
    }
    catch {
        Write-Host "Error scanning $folderPath for executables: $_" -ForegroundColor Red
    }
    
    Write-Progress -Id 2 -Activity "Finding executables" -Completed
    return $exeFiles
}

# ====================================================================================
# MAIN PROCESSING LOOP - Process each game folder
# ====================================================================================

# Get list of game folders to process (immediate subdirectories only)
Write-Host "Checking for game folders..." -ForegroundColor Yellow
$gameFolders = Get-ChildItem -Path $rootGameFolder -Directory

# Validate that we found some folders to process
if ($gameFolders.Count -eq 0) {
    Write-Host "No subdirectories found in $rootGameFolder" -ForegroundColor Red
    Write-Host "Make sure the path contains game folders to scan." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

# Initialize counters for progress tracking and final summary
$total = $gameFolders.Count
$index = 0
$created = 0        # Successfully created shortcuts
$skipped = 0        # Shortcuts already existed or duplicates
$notFound = 0       # No suitable executable found
$sanitized = 0      # Created with sanitized/safe filename
$errors = 0         # Errors during processing
$timeouts = 0       # Folders that hit the timeout limit

# Start timing for performance tracking
$startTime = Get-Date
$lastUpdateTime = $startTime

Write-Host "Starting to process $total game folders..." -ForegroundColor Green

# Process each game folder
foreach ($folder in $gameFolders) {
    $index++
    $gameName = $folder.Name
    
    # Calculate timing statistics for progress estimation
    $currentTime = Get-Date
    $elapsedTime = $currentTime - $startTime
    $itemsRemaining = $total - $index
    
    # Only recalculate estimated time every 5 items or every 10 seconds to avoid fluctuations
    # This prevents the time estimate from jumping around too much
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
    
    # Update the main progress bar
    Write-Progress -Id 1 -Activity "Scanning Games" -Status $progressStatus -PercentComplete (($index / $total) * 100)
    
    # Special handling for known problematic games
    # Some games have known issues with the standard search, so handle them specially
    $manualExePath = $null
    if ($gameName -eq "ELDEN RING") {
        # Elden Ring has a known deep structure that our search might miss
        $specificPath = "$($folder.FullName)\Game\eldenring.exe"
        if (Test-Path $specificPath) {
            Write-Host "Found Elden Ring executable via direct path!" -ForegroundColor Green
            $manualExePath = $specificPath
        }
    }
    
    # Use our optimized function to find executables with improved recursive search
    try {
        $exeFiles = @()
        
        # Use manual path if we have one for special cases
        if ($manualExePath) {
            $exeFiles = @(Get-Item -Path $manualExePath)
        } else {
            # Use our smart search function
            $searchStartTime = Get-Date
            $exeFiles = Find-Executables -folderPath $folder.FullName -maxDepth $maxDepth -gameName $gameName -timeout $folderTimeout
            $searchTime = (Get-Date) - $searchStartTime
            
            # Check if we likely hit a timeout (over 90% of the timeout time used)
            # This helps identify problematic folders for debugging
            if ($searchTime.TotalSeconds -gt ($folderTimeout * 0.9)) {
                Write-Host "Warning: $gameName search took $($searchTime.TotalSeconds) seconds" -ForegroundColor Yellow
                $timeouts++
                # Log the timeout to errors log for analysis
                "$gameName - Timeout reached while scanning for executables" | Out-File -FilePath $logErrors -Append
            }
        }
        
        # If no executables found at all, log and continue
        if ($exeFiles.Count -eq 0) {
            "$gameName - No executable found" | Out-File -FilePath $logNotFound -Append
            $notFound++
            continue
        }
        
        # Score all found executables and filter out blocked ones
        $scored = $exeFiles | ForEach-Object {
            [PSCustomObject]@{
                Path  = $_.FullName
                Score = Get-ExecutableScore -exePath $_.FullName -gameFolderName $gameName
            }
        } | Where-Object { $_.Score -ge 0 } | Sort-Object Score -Descending  # Sort by score, highest first
        
        # If all executables were filtered out (negative scores), log and continue
        if ($scored.Count -eq 0) {
            "$gameName - No suitable exe (all files blocked by filters)" | Out-File -FilePath $logNotFound -Append
            # Also log what files were found but blocked for debugging
            foreach ($exe in $exeFiles) {
                $exeName = [System.IO.Path]::GetFileNameWithoutExtension($exe.FullName).ToLower()
                "$gameName - Found but filtered: $($exe.FullName) (score would be negative)" | Out-File -FilePath $logBlocked -Append
            }
            $notFound++
            continue
        }
        
        # Choose the highest-scoring executable
        $chosenExe = $scored[0].Path
        
        # Attempt to create the shortcut
        $shortcutStatus = Create-Shortcut -targetExe $chosenExe -shortcutName $gameName -outputFolder $shortcutOutputFolder
        
        # Log the result and update counters based on what happened
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
        # Log any unexpected errors and continue with the next folder
        Write-Host "Error processing $gameName`: $_" -ForegroundColor Red
        "$gameName - Error: $_" | Out-File -FilePath $logErrors -Append
        $errors++
    }
}

# ====================================================================================
# COMPLETION AND SUMMARY - Show final results
# ====================================================================================

# Calculate total processing time
$totalTime = (Get-Date) - $startTime
$formattedTotalTime = Format-TimeSpan -TimeSpan $totalTime

# Clear the progress bar
Write-Progress -Id 1 -Activity "Scanning Games" -Completed

# Display completion summary
Write-Host ""
Write-Host "==================== SCAN COMPLETE ====================" -ForegroundColor Green
Write-Host "Done! Shortcuts saved to: $shortcutOutputFolder" -ForegroundColor Green
Write-Host "Total time: $formattedTotalTime" -ForegroundColor Cyan
Write-Host "Results: $created created, $sanitized sanitized, $skipped skipped, $notFound not found, $errors errors, $timeouts timeouts" -ForegroundColor White
Write-Host "See FoundGames.log, NotFoundGames.log, SkippedGames.log, ErrorGames.log, and BlockedFiles.log for details" -ForegroundColor Yellow
Write-Host ""

# Pause so user can see results before window closes
Read-Host "Press Enter to exit"
