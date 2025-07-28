# Checks and installs prerequisites (CPU‑Z) if needed
function Install-CpuZ {
    $cpuZPath = "C:\Program Files\CPUID\CPU-Z\cpuz.exe"
    
    # Fast path: CPU-Z is already installed in the default location
    if (Test-Path $cpuZPath) {
        return $true
    }
    
    # Admin rights are required for installing software
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Error "Administrator privileges are required to install CPU-Z"
        return $false
    }
    
    # Try Winget first
    try {
        Write-Host "Attempting to install CPU-Z via Winget..."
        winget install --id CPUID.CPU-Z -e --source winget -h
        if (Test-Path $cpuZPath) {
            return $true
        }
    } catch {
        Write-Warning "Winget is not available, falling back to Chocolatey..."
    }
    
    # Install Chocolatey if missing
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Host "Installing Chocolatey..."
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Refresh-Environment
    }
    
    # Install CPU-Z via Chocolatey
    try {
        Write-Host "Installing CPU-Z via Chocolatey..."
        choco install cpu-z -y --force
        return $true
    } catch {
        Write-Error "Failed to install CPU-Z: $_"
        return $false
    }
}

# Gathers CPU information by running CPU-Z and parsing its TXT report
function Get-CpuInfoFromCpuZ {
    if (-not (Install-CpuZ)) {
        Write-Error "CPU-Z is not installed and cannot be installed automatically"
        return $null
    }
    
    $cpuZPath = "C:\Program Files\CPUID\CPU-Z\cpuz.exe"
    $reportName = "cpuz_report"
    $reportPath = Join-Path $env:TEMP $reportName
    
    # Remove any previous reports
    if (Test-Path "$reportPath.txt") { Remove-Item "$reportPath.txt" -Force }
    if (Test-Path "$reportPath.TXT") { Remove-Item "$reportPath.TXT" -Force }
    
    # Run CPU-Z to generate the report
    $process = Start-Process -FilePath $cpuZPath -ArgumentList "-txt=$reportPath" -Wait -PassThru -WindowStyle Hidden
    Start-Sleep -Seconds 2
    
    # Locate the resulting report (case-insensitive extension)
    $finalReportPath = if (Test-Path "$reportPath.txt") { "$reportPath.txt" }
    elseif (Test-Path "$reportPath.TXT") { "$reportPath.TXT" }
    else {
        Write-Error "Failed to create CPU-Z report"
        return $null
    }
    
    $reportContent = Get-Content $finalReportPath -Raw
    
    # === CPU parsing (unchanged logic) ===
    $result = [PSCustomObject]@{
        Sockets    = 0
        P_Cores    = 0
        E_Cores    = 0
        Threads    = 0
        Name       = ""
        BaseFreq   = ""
        MaxFreq    = ""
        SocketType = ""
    }
    
    # Global counts
    if ($reportContent -match "Number of sockets\s+(\d+)") {
        $result.Sockets = [int]$matches[1]
    }
    if ($reportContent -match "Number of threads\s+(\d+)") {
        $result.Threads = [int]$matches[1]
    }
    
    # Extract the first "Socket N" CPU info block
    if ($reportContent -match "Processors Information[\s-]+(Socket \d+\s+ID = \d+[\s\S]+?)(?=\n\s{2,}\S+|\z)") {
        $cpuInfo = $matches[1]
        
        # CPU Name
        if ($cpuInfo -match "Name\s+(.+)") {
            $result.Name = $matches[1].Trim()
        }
        # Socket type (“Socket 1700 LGA” → “1700 LGA”, “Socket AM5 (LGA1718)” → “AM5”)
        if ($cpuInfo -match 'Package(?:\s*\([^)]+\))?\s+Socket\s+([^(\r\n]+)') {
            $result.SocketType = $matches[1].Trim()
        }
        # Hybrid (Intel P/E-cores)
        if ($cpuInfo -match "Core Set 0\s+P-Cores, (\d+) cores") {
            $result.P_Cores = [int]$matches[1]
        }
        if ($cpuInfo -match "Core Set 1\s+E-Cores, (\d+) cores") {
            $result.E_Cores = [int]$matches[1]
        }
        # Non-hybrid (AMD/Xeon) — do not divide by sockets
        if ($cpuInfo -match "Number of cores\s+(\d+)") {
            $cores = [int]$matches[1]
            if ($result.P_Cores -eq 0 -and $result.E_Cores -eq 0) {
                $result.P_Cores = $cores
            }
        }
        # Frequencies (marketing/base/max)
        if ($cpuInfo -match "(?:Stock frequency|Base frequency \(cores\))\s+(\d+) MHz") {
            $baseFreq = [decimal]$matches[1] / 1000
            $result.BaseFreq = $baseFreq.ToString("0.0").Replace(".", ",")
        }
        if ($cpuInfo -match "Max frequency\s+(\d+) MHz") {
            $maxFreq = [decimal]$matches[1] / 1000
            $result.MaxFreq = $maxFreq.ToString("0.0").Replace(".", ",")
        }
    }

    # === MEMORY parsing (more robust; avoids false "mixed") ===
    $memoryInfo = [PSCustomObject]@{
        TotalGB     = 0           # Sum of installed module capacities
        Config      = ""          # e.g., "4×48 GB DDR5-5600 MHz" or "... (mixed)"
        TotalSlots  = 0           # Across all DMI Physical Memory Arrays
        UsedSlots   = 0           # Count of installed modules
        RawModules  = @()         # Per-module raw info
    }

    # Capture all "DMI Memory Device" blocks (tolerant to whitespace)
    $memoryBlocks = [regex]::Matches(
        $reportContent,
        '(?s)^DMI Memory Device\s*\r?\n(.*?)(?:\r?\n\r?\n)',
        [System.Text.RegularExpressions.RegexOptions]::Multiline
    )
    
    $modules     = @()
    $totalSizeGB = 0
    $usedSlots   = 0

    foreach ($m in $memoryBlocks) {
        $blockText = $m.Groups[1].Value

        $module = [ordered]@{
            SizeGB      = 0
            Type        = ''
            SpeedMHz    = 0
            IsInstalled = $false
        }

        # Installed module has an explicit numeric "size <N> GB"
        if ($blockText -match '(?mi)^\s*size\s+(\d+)\s*GB') {
            $module.SizeGB   = [int]$matches[1]
            $module.IsInstalled = $true
            $totalSizeGB     += $module.SizeGB
            $usedSlots++
        }

        # Normalize type: trim and uppercase (e.g., "DDR5")
        if ($blockText -match '(?mi)^\s*type\s+([^\r\n]+)') {
            $module.Type = ($matches[1] -replace '\s+',' ').Trim().ToUpperInvariant()
        }

        # Speed: integer in MHz; keep 0 if not present
        if ($blockText -match '(?mi)^\s*speed\s+(\d+)\s*MHz') {
            $module.SpeedMHz = [int]$matches[1]
        }

        $modules += [pscustomobject]$module
    }

    # Sum "max# of devices" across all "DMI Physical Memory Array" blocks
    $arrayMatches = [regex]::Matches(
        $reportContent,
        '(?mis)^DMI Physical Memory Array\s+.*?^\s*max# of devices\s+(\d+)',
        [System.Text.RegularExpressions.RegexOptions]::Multiline
    )
    $memoryInfo.TotalSlots = (
        $arrayMatches |
        ForEach-Object { [int]$_.Groups[1].Value } |
        Measure-Object -Sum
    ).Sum

    # Build configuration string based on installed modules
    $installed = $modules | Where-Object { $_.IsInstalled }

    if ($installed.Count -gt 0) {
        # Unique groups (ignore zero speeds to avoid false "mixed")
        $sizesUnique  = $installed | Select-Object -ExpandProperty SizeGB     | Sort-Object -Unique
        $typesUnique  = $installed | Select-Object -ExpandProperty Type       | Sort-Object -Unique
        $speedsUnique = $installed | Where-Object { $_.SpeedMHz -gt 0 } | Select-Object -ExpandProperty SpeedMHz | Sort-Object -Unique

        $allSameSize  = $sizesUnique.Count  -eq 1
        $allSameType  = $typesUnique.Count  -eq 1 -and $typesUnique[0]
        $allSameSpeed = $speedsUnique.Count -le 1  # <=1 means all equal or all missing/zero

        # Representative type/speed
        $repType  = $typesUnique | Select-Object -First 1
        $repSpeed = $speedsUnique | Select-Object -First 1

        if ($allSameSize -and $allSameType -and $allSameSpeed) {
            # Example: "4×48 GB DDR5-3600 MHz"
            $memoryInfo.Config = ("{0}×{1} GB {2}{3}" -f
                $installed.Count,
                $sizesUnique[0],
                $repType,
                ($(if ($repSpeed) { "-$repSpeed MHz" } else { "" }))
            ).Trim()
        }
        else {
            # Mixed sizes/types/speeds → show grouped sizes and tag as mixed
            $sizeGroups = $installed | Group-Object SizeGB | ForEach-Object { "{0}×{1} GB" -f $_.Count, $_.Name }
            $suffixType = if ($repType) { " $repType" } else { "" }
            $memoryInfo.Config = (($sizeGroups -join " + ") + "$suffixType (mixed)").Trim()
        }

        # ECC tag if any array reports ECC correction
        $hasECC = [regex]::IsMatch($reportContent, '(?mi)^\s*DMI Physical Memory Array\s+.*^\s*correction\s+.*ECC')
        if ($hasECC) {
            $memoryInfo.Config += " ECC"
        }
    }

    # Finalize memory info object
    $memoryInfo.TotalGB   = [int]$totalSizeGB
    $memoryInfo.UsedSlots = [int]$usedSlots
    $memoryInfo.RawModules= $modules

    # Attach memory info to the main result
    $result | Add-Member -NotePropertyName "MemoryInfo" -NotePropertyValue $memoryInfo

    # Cleanup temp report
    Remove-Item $finalReportPath -Force -ErrorAction SilentlyContinue
    
    return $result
}

# Refreshes current PowerShell process environment variables from registry
function Refresh-Environment {
    foreach($level in "Machine","User") {
        [Environment]::GetEnvironmentVariables($level).GetEnumerator() | ForEach-Object {
            if ($_.Name -match '^(?:path|pathext)$') {
                $_.Value = $_.Value -replace ';+',';'
                $_.Value = $_.Value.Trim(';')
            }
            Set-Content "env:\$($_.Name)" $_.Value
        }
    }
}

# Main: run, format, and print
try {
    $cpuInfo = Get-CpuInfoFromCpuZ
    
    if ($cpuInfo) {
        # CPU cores string (e.g., "2x28P+0E" or "2x28")
        $coreInfo = if ($cpuInfo.E_Cores -gt 0) {
            "$($cpuInfo.Sockets)x$($cpuInfo.P_Cores)P+$($cpuInfo.E_Cores)E"
        } else {
            "$($cpuInfo.Sockets)x$($cpuInfo.P_Cores)"
        }
        
        # CPU frequency string (e.g., "3,4/5,6GHz" or "3,4GHz")
        $freqInfo = if ($cpuInfo.BaseFreq -and $cpuInfo.MaxFreq) {
            "$($cpuInfo.BaseFreq)/$($cpuInfo.MaxFreq)GHz"
        } elseif ($cpuInfo.BaseFreq) {
            "$($cpuInfo.BaseFreq)GHz"
        } elseif ($cpuInfo.MaxFreq) {
            "?/$($cpuInfo.MaxFreq)GHz"
        } else {
            ""
        }
        
        # CPU line
        "$($cpuInfo.Name.Trim()) | $coreInfo $($cpuInfo.Threads) | $freqInfo | $($cpuInfo.SocketType)"
        
        # Memory line
        $mem = $cpuInfo.MemoryInfo
        if ($mem.TotalSlots -gt 0) {
            "$($mem.TotalGB) GB ($($mem.Config)) in $($mem.TotalSlots) slots, $($mem.UsedSlots) used"
        }
        else {
            "$($mem.TotalGB) GB ($($mem.Config))"
        }
    }
} catch {
    Write-Error "Error: $_"
}
