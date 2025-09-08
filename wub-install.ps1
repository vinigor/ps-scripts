# Windows Update Blocker Download and PATH Setup Script

param(
    [string]$DownloadUrl = "https://www.dropbox.com/scl/fi/e787say9wv6sa0e1x16dd/Wub.exe?rlkey=8bfu4n348g2bu0riu1rqbl8i6&st=2coxl61e&dl=1",
    [string]$DestinationPath = "C:\Users\admin\Desktop\Wub",
    [string]$FileName = "Wub.exe"
)

function Write-Status {
    param([string]$Message, [string]$Color = "Green")
    Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message" -ForegroundColor $Color
}

function Test-AdminRights {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Show-Menu {
    param([string]$Title, [string[]]$Options)
    
    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "[$($i + 1)] $($Options[$i])" -ForegroundColor Yellow
    }
    Write-Host "[0] Skip" -ForegroundColor Gray
    Write-Host ""
}

function Get-UserChoice {
    param([int]$MaxOption)
    
    do {
        $choice = Read-Host "Select option (0-$MaxOption)"
        if ($choice -match '^\d+$' -and [int]$choice -ge 0 -and [int]$choice -le $MaxOption) {
            return [int]$choice
        }
        Write-Host "Invalid choice! Enter number from 0 to $MaxOption" -ForegroundColor Red
    } while ($true)
}

function Download-File {
    param([string]$Url, [string]$FilePath)
    
    try {
        Write-Status "Downloading file..."
        Write-Status "URL: $Url" "Cyan"
        Write-Status "Destination: $FilePath" "Cyan"
        
        # Create WebClient with settings
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        
        # Download file
        $webClient.DownloadFile($Url, $FilePath)
        $webClient.Dispose()
        
        # Check download success
        if (Test-Path $FilePath) {
            $fileSize = (Get-Item $FilePath).Length
            Write-Status "File downloaded successfully! Size: $([Math]::Round($fileSize/1KB, 2)) KB"
            return $true
        } else {
            Write-Status "Error: file not found after download" "Red"
            return $false
        }
    }
    catch {
        Write-Status "Download error: $($_.Exception.Message)" "Red"
        
        # Alternative method via Invoke-WebRequest
        try {
            Write-Status "Trying alternative download method..." "Yellow"
            Invoke-WebRequest -Uri $Url -OutFile $FilePath -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            
            if (Test-Path $FilePath) {
                Write-Status "File downloaded using alternative method!"
                return $true
            }
        }
        catch {
            Write-Status "Alternative method also failed: $($_.Exception.Message)" "Red"
        }
        
        return $false
    }
}

function Add-ToSystemAndSessionPath {
    param([string]$PathToAdd)
    
    try {
        Write-Status "Adding to system PATH (all users)..."
        [Environment]::SetEnvironmentVariable("PATH", $env:PATH + ";$PathToAdd", [EnvironmentVariableTarget]::Machine)
        
        Write-Status "Adding to current session PATH..."
        $env:PATH += ";$PathToAdd"
        
        Write-Status "Path added successfully to both system and current session!" "Green"
        return $true
    }
    catch {
        Write-Status "Error adding to PATH: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Main logic
Write-Status "=== Windows Update Blocker Downloader ===" "Cyan"

# Create target folder if it doesn't exist
$fullFilePath = Join-Path $DestinationPath $FileName

try {
    if (-not (Test-Path $DestinationPath)) {
        Write-Status "Creating folder: $DestinationPath"
        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    }
}
catch {
    Write-Status "Error creating folder: $($_.Exception.Message)" "Red"
    exit 1
}

# Check if file exists
if (Test-Path $fullFilePath) {
    Write-Status "File already exists: $fullFilePath" "Yellow"
    
    $fileInfo = Get-Item $fullFilePath
    Write-Status "File size: $([Math]::Round($fileInfo.Length/1KB, 2)) KB"
    Write-Status "Creation date: $($fileInfo.CreationTime)"
    
    $overwrite = Read-Host "Overwrite file? (y/N)"
    if ($overwrite -notmatch '^[yY]') {
        Write-Status "Download skipped"
        $downloadSuccess = $true
    } else {
        $downloadSuccess = Download-File -Url $DownloadUrl -FilePath $fullFilePath
    }
} else {
    Write-Status "File not found, starting download..."
    $downloadSuccess = Download-File -Url $DownloadUrl -FilePath $fullFilePath
}

if (-not $downloadSuccess) {
    Write-Status "Failed to download file. Exiting." "Red"
    exit 1
}

# PATH configuration
Write-Status ""
Write-Status "=== PATH Variable Configuration ===" "Cyan"

# Check current PATH
$currentPath = $env:PATH
$targetPath = $DestinationPath

if ($currentPath -like "*$targetPath*") {
    Write-Status "Path already exists in PATH variable" "Yellow"
} else {
    Write-Status "Path not found in PATH variable"
    
    # Check admin rights
    $isAdmin = Test-AdminRights
    if ($isAdmin) {
        Write-Status "Administrator rights detected" "Green"
    } else {
        Write-Status "Administrator rights not detected" "Yellow"
        Write-Status "Administrator rights required to modify system PATH variable" "Yellow"
    }
    
    # Menu options
    $menuOptions = @()
    
    if ($isAdmin) {
        $menuOptions += "Add to system PATH (all users) + current session"
        $menuOptions += "Show commands for manual execution"
    } else {
        $menuOptions += "Add to user PATH (current user only)"
        $menuOptions += "Add to current session only"
        $menuOptions += "Show commands for manual execution"
    }
    
    Show-Menu -Title "Choose PATH configuration method" -Options $menuOptions
    $choice = Get-UserChoice -MaxOption $menuOptions.Count
    
    if ($choice -eq 0) {
        Write-Status "PATH configuration skipped"
    } elseif (($isAdmin -and $choice -eq 2) -or (-not $isAdmin -and $choice -eq 3)) {
        # Show commands for manual execution
        Write-Status ""
        Write-Status "=== Commands for Manual Execution ===" "Cyan"
        Write-Status ""
        
        Write-Status "System PATH (requires administrator rights):" "Yellow"
        Write-Host '[Environment]::SetEnvironmentVariable("PATH", $env:PATH + ";C:\Users\admin\Desktop\Wub", [EnvironmentVariableTarget]::Machine)' -ForegroundColor White
        Write-Status ""
        
        Write-Status "Current session PATH:" "Yellow"
        Write-Host '$env:PATH += ";C:\Users\admin\Desktop\Wub"' -ForegroundColor White
        Write-Status ""
        
        Write-Status "Combined command (system + session):" "Yellow"
        Write-Host '[Environment]::SetEnvironmentVariable("PATH", $env:PATH + ";C:\Users\admin\Desktop\Wub", [EnvironmentVariableTarget]::Machine); $env:PATH += ";C:\Users\admin\Desktop\Wub"' -ForegroundColor White
        Write-Status ""
        
    } else {
        # Execute selected option
        $confirm = Read-Host "Confirm execution? (Y/n)"
        if ($confirm -notmatch '^[nN]') {
            try {
                if ($isAdmin -and $choice -eq 1) {
                    # Add to system PATH + current session
                    $success = Add-ToSystemAndSessionPath -PathToAdd $targetPath
                } elseif (-not $isAdmin -and $choice -eq 1) {
                    # Add to user PATH only
                    Write-Status "Adding to user PATH..."
                    [Environment]::SetEnvironmentVariable("PATH", $env:PATH + ";$targetPath", [EnvironmentVariableTarget]::User)
                    Write-Status "Added to user PATH successfully!" "Green"
                    $success = $true
                } elseif (-not $isAdmin -and $choice -eq 2) {
                    # Add to current session only
                    Write-Status "Adding to current session PATH..."
                    $env:PATH += ";$targetPath"
                    Write-Status "Added to current session PATH!" "Green"
                    $success = $true
                }
                
                if ($success) {
                    Write-Status "Checking Wub command availability..."
                    try {
                        $wubTest = & "$fullFilePath" 2>$null
                        Write-Status "Wub command is available!" "Green"
                    }
                    catch {
                        Write-Status "Wub command not yet available. May require PowerShell restart." "Yellow"
                    }
                }
                
            }
            catch {
                Write-Status "Error executing command: $($_.Exception.Message)" "Red"
            }
        } else {
            Write-Status "Command execution cancelled"
        }
    }
}

Write-Status ""
Write-Status "=== Setup Complete ===" "Cyan"
Write-Status "Windows Update Blocker file location: $fullFilePath"

if (Test-Path $fullFilePath) {
    Write-Status "File size: $([Math]::Round((Get-Item $fullFilePath).Length/1KB, 2)) KB"
}

Write-Status ""
Write-Status "Usage commands:" "Yellow"
Write-Status "wub /D  - disable Windows updates" "Cyan"
Write-Status "wub /E  - enable Windows updates" "Cyan"
Write-Status ""

Read-Host "Press Enter to exit"
