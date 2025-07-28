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
    
    # Try Winget first (left as-is)
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

# Generates a single CPU-Z TXT report and returns its content.
# Shows a lightweight progress spinner while waiting for the file to appear.
function Invoke-CpuZReport {
    if (-not (Install-CpuZ)) {
        throw "CPU-Z is not installed and cannot be installed automatically"
    }

    $cpuZPath    = "C:\Program Files\CPUID\CPU-Z\cpuz.exe"
    $reportBase  = Join-Path $env:TEMP ("cpuz_report_" + ([guid]::NewGuid().Guid))
    $lowerPath   = "$reportBase.txt"
    $upperPath   = "$reportBase.TXT"

    if (Test-Path $lowerPath) { Remove-Item $lowerPath -Force -ErrorAction SilentlyContinue }
    if (Test-Path $upperPath) { Remove-Item $upperPath -Force -ErrorAction SilentlyContinue }

    # Kick off CPU-Z report generation
    Start-Process -FilePath $cpuZPath -ArgumentList "-txt=$reportBase" -WindowStyle Hidden | Out-Null

    # Wait for the file to appear (polling, no hard sleep)
    $spinner = @('|','/','-','\')
    $idx = 0
    $timeoutSec = 12
    $start = Get-Date

    while (-not (Test-Path $lowerPath) -and -not (Test-Path $upperPath)) {
        $elapsed = (Get-Date) - $start
        if ($elapsed.TotalSeconds -ge $timeoutSec) {
            Write-Host "`rGenerating CPU-Z report... timeout after $timeoutSec s" 
            throw "CPU-Z report was not created in time"
        }
        Write-Host -NoNewline ("`rGenerating CPU-Z report... " + $spinner[$idx % $spinner.Length])
        $idx++
        Start-Sleep -Milliseconds 150
    }
    Write-Host ("`rGenerating CPU-Z report... done".PadRight(60))

    $final = if (Test-Path $lowerPath) { $lowerPath } else { $upperPath }
    $content = Get-Content $final -Raw

    # Cleanup
    Remove-Item $final -Force -ErrorAction SilentlyContinue
    return $content
}

# Parses CPU and Memory from a given CPU-Z report (unchanged logic, just parameterized)
function Get-CpuInfoFromCpuZ {
    param(
        [string]$ReportContent  # When provided, we parse it directly (no re-run of CPU-Z)
    )

    # If report is not provided, fall back to legacy behavior (single-run inside)
    if (-not $ReportContent) {
        # Legacy path (kept as-is)
        $cpuZPath = "C:\Program Files\CPUID\CPU-Z\cpuz.exe"
        if (-not (Install-CpuZ)) {
            Write-Error "CPU-Z is not installed and cannot be installed automatically"
            return $null
        }
        $reportName = "cpuz_report"
        $reportPath = Join-Path $env:TEMP $reportName
        if (Test-Path "$reportPath.txt") { Remove-Item "$reportPath.txt" -Force }
        if (Test-Path "$reportPath.TXT") { Remove-Item "$reportPath.TXT" -Force }
        $process = Start-Process -FilePath $cpuZPath -ArgumentList "-txt=$reportPath" -Wait -PassThru -WindowStyle Hidden
        Start-Sleep -Seconds 2
        $finalReportPath = if (Test-Path "$reportPath.txt") { "$reportPath.txt" }
                           elseif (Test-Path "$reportPath.TXT") { "$reportPath.TXT" }
                           else { Write-Error "Failed to create CPU-Z report"; return $null }
        $ReportContent = Get-Content $finalReportPath -Raw
        Remove-Item $finalReportPath -Force -ErrorAction SilentlyContinue
    }

    $reportContent = $ReportContent

    # === CPU parsing (unchanged) ===
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
    
    if ($reportContent -match "Number of sockets\s+(\d+)") {
        $result.Sockets = [int]$matches[1]
    }
    if ($reportContent -match "Number of threads\s+(\d+)") {
        $result.Threads = [int]$matches[1]
    }
    if ($reportContent -match "Processors Information[\s-]+(Socket \d+\s+ID = \d+[\s\S]+?)(?=\n\s{2,}\S+|\z)") {
        $cpuInfo = $matches[1]
        if ($cpuInfo -match "Name\s+(.+)") {
            $result.Name = $matches[1].Trim()
        }
        if ($cpuInfo -match 'Package(?:\s*\([^)]+\))?\s+Socket\s+([^(\r\n]+)') {
            $result.SocketType = $matches[1].Trim()
        }
        if ($cpuInfo -match "Core Set 0\s+P-Cores, (\d+) cores") {
            $result.P_Cores = [int]$matches[1]
        }
        if ($cpuInfo -match "Core Set 1\s+E-Cores, (\d+) cores") {
            $result.E_Cores = [int]$matches[1]
        }
        if ($cpuInfo -match "Number of cores\s+(\d+)") {
            $cores = [int]$matches[1]
            if ($result.P_Cores -eq 0 -and $result.E_Cores -eq 0) {
                $result.P_Cores = $cores
            }
        }
        if ($cpuInfo -match "(?:Stock frequency|Base frequency \(cores\))\s+(\d+) MHz") {
            $baseFreq = [decimal]$matches[1] / 1000
            $result.BaseFreq = $baseFreq.ToString("0.0").Replace(".", ",")
        }
        if ($cpuInfo -match "Max frequency\s+(\d+) MHz") {
            $maxFreq = [decimal]$matches[1] / 1000
            $result.MaxFreq = $maxFreq.ToString("0.0").Replace(".", ",")
        }
    }

    # === MEMORY parsing (robust; avoids false "mixed") ===
    $memoryInfo = [PSCustomObject]@{
        TotalGB     = 0
        Config      = ""
        TotalSlots  = 0
        UsedSlots   = 0
        RawModules  = @()
    }

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
        if ($blockText -match '(?mi)^\s*size\s+(\d+)\s*GB') {
            $module.SizeGB   = [int]$matches[1]
            $module.IsInstalled = $true
            $totalSizeGB     += $module.SizeGB
            $usedSlots++
        }
        if ($blockText -match '(?mi)^\s*type\s+([^\r\n]+)') {
            $module.Type = ($matches[1] -replace '\s+',' ').Trim().ToUpperInvariant()
        }
        if ($blockText -match '(?mi)^\s*speed\s+(\d+)\s*MHz') {
            $module.SpeedMHz = [int]$matches[1]
        }
        $modules += [pscustomobject]$module
    }

    $arrayMatches = [regex]::Matches(
        $reportContent,
        '(?mis)^DMI Physical Memory Array\s+.*?^\s*max# of devices\s+(\d+)',
        [System.Text.RegularExpressions.RegexOptions]::Multiline
    )
    $totalSlots = (
        $arrayMatches |
        ForEach-Object { [int]$_.Groups[1].Value } |
        Measure-Object -Sum
    ).Sum

    $installed = $modules | Where-Object { $_.IsInstalled }
    $config = ""
    if ($installed.Count -gt 0) {
        $sizesUnique  = $installed | Select-Object -ExpandProperty SizeGB     | Sort-Object -Unique
        $typesUnique  = $installed | Select-Object -ExpandProperty Type       | Sort-Object -Unique
        $speedsUnique = $installed | Where-Object { $_.SpeedMHz -gt 0 } | Select-Object -ExpandProperty SpeedMHz | Sort-Object -Unique

        $allSameSize  = $sizesUnique.Count  -eq 1
        $allSameType  = $typesUnique.Count  -eq 1 -and $typesUnique[0]
        $allSameSpeed = $speedsUnique.Count -le 1  # <=1 means equal or missing

        $repType  = $typesUnique  | Select-Object -First 1
        $repSpeed = $speedsUnique | Select-Object -First 1

        if ($allSameSize -and $allSameType -and $allSameSpeed) {
            $config = ("{0}×{1} GB {2}{3}" -f
                $installed.Count,
                $sizesUnique[0],
                $repType,
                ($(if ($repSpeed) { "-$repSpeed MHz" } else { "" }))
            ).Trim()
        } else {
            $sizeGroups = $installed | Group-Object SizeGB | ForEach-Object { "{0}×{1} GB" -f $_.Count, $_.Name }
            $suffixType = if ($repType) { " $repType" } else { "" }
            $config = (($sizeGroups -join " + ") + "$suffixType (mixed)").Trim()
        }

        $hasECC = [regex]::IsMatch($reportContent, '(?mi)^\s*DMI Physical Memory Array\s+.*^\s*correction\s+.*ECC')
        if ($hasECC) { $config += " ECC" }
    }

    $memoryInfo.TotalGB    = [int]$totalSizeGB
    $memoryInfo.UsedSlots  = [int]$usedSlots
    $memoryInfo.TotalSlots = [int]$totalSlots
    $memoryInfo.Config     = $config
    $memoryInfo.RawModules = $modules

    $result | Add-Member -NotePropertyName "MemoryInfo" -NotePropertyValue $memoryInfo
    return $result
}

# Parses "Drive ..." blocks from CPU-Z report and returns ONLY virtual drives
function Get-VirtualDrivesFromCpuZ {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ReportContent
    )

    $virtual = @()
    if (-not $ReportContent) { return $virtual }

    $driveMatches = [regex]::Matches(
        $ReportContent,
        '(?ms)^Drive\s+[^\r\n]+\r?\n(.*?)(?=^Drive\s+|\z)',
        [System.Text.RegularExpressions.RegexOptions]::Multiline
    )

    foreach ($m in $driveMatches) {
        $block = $m.Groups[1].Value

        $name = ''
        $capacityGB = $null
        $bus = ''
        $dtype = ''

        if ($block -match '(?mi)^\s*Name\s+([^\r\n]+)') { $name = $matches[1].Trim() }
        if ($block -match '(?mi)^\s*Capacity\s+([\d\.,]+)\s*GB') {
            $capStr = $matches[1].Replace(',', '.')
            $capacityGB = [double]::Parse($capStr, [System.Globalization.CultureInfo]::InvariantCulture)
        }
        if ($block -match '(?mi)^\s*Bus Type\s+([^\r\n]+)') { $bus = $matches[1].Trim() }
        if ($block -match '(?mi)^\s*Type\s+([^\r\n]+)')     { $dtype = $matches[1].Trim() }

        $isVirtual = ($name -match '(?i)\b(QEMU|VBOX|VMWARE|VIRTIO|HYPER-V|KVM|MSFT VIRTUAL|VIRTUAL)\b')
        if ($isVirtual) {
            $sizeStr = if ($capacityGB -ge 1000) { ("{0:N1} TB" -f ($capacityGB/1000)) } else { ("{0:N1} GB" -f $capacityGB) }
            $virtual += [pscustomobject]@{
                IsVirtual = $true
                MediaType = 'VIRTUAL'
                Model     = $name
                Size      = $sizeStr
                BusType   = $bus
            }
        }
    }
    return $virtual
}

# Builds a list of physical disks from Get-PhysicalDisk, excluding obvious virtual models
function Get-PhysicalDrivesList {
    $list = @()
    try {
        $pd = Get-PhysicalDisk
    } catch {
        return $list
    }

    foreach ($d in $pd) {
        $model = if ($d.PSObject.Properties.Name -contains 'Model' -and $d.Model) { $d.Model }
                 elseif ($d.FriendlyName) { $d.FriendlyName }
                 else { '' }

        # Skip items that look virtual to avoid duplicating CPU-Z-derived entries
        $looksVirtual = ($model -match '(?i)\b(QEMU|VBOX|VMWARE|VIRTIO|HYPER-V|KVM|MSFT VIRTUAL|VIRTUAL)\b')
        if ($looksVirtual) { continue }

        $type  = if ($d.MediaType) { $d.MediaType } else { 'Unspecified' }
        $bytes = $d.Size

        $sizeStr = if ($bytes -ge 1e12) { "{0:N1} TB" -f ($bytes/1e12) } else { "{0:N1} GB" -f ($bytes/1e9) }

        $list += [pscustomobject]@{
            IsVirtual = $false
            MediaType = $type
            Model     = $model
            Size      = $sizeStr
        }
    }
    return $list
}

# Extracts Windows Version from CPU-Z report; falls back to registry when absent
function Get-WindowsVersionString {
    param([string]$ReportContent)
    if ($ReportContent -and ($ReportContent -match '(?mi)^\s*Windows\s+Version\s+(.+)$')) {
        return $matches[1].Trim()
    }
    try {
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $product = $cv.ProductName
        $display = $cv.DisplayVersion
        if (-not $display) { $display = $cv.ReleaseId }
        $build   = $cv.CurrentBuild
        $ubr     = $cv.UBR
        $buildStr = if ($ubr -ge 0) { "$build.$ubr" } else { "$build" }
        return "$product $display ($buildStr)"
    } catch {
        return ""
    }
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

# -------------------------
# Main: single CPU-Z run
# -------------------------
try {
    # Single report for everything (CPU, RAM, virtual drives, Windows version)
    $cpuzRaw = Invoke-CpuZReport

    # CPU + RAM
    $cpuInfo = Get-CpuInfoFromCpuZ -ReportContent $cpuzRaw
    if ($cpuInfo) {
        $coreInfo = if ($cpuInfo.E_Cores -gt 0) { "$($cpuInfo.Sockets)x$($cpuInfo.P_Cores)P+$($cpuInfo.E_Cores)E" } else { "$($cpuInfo.Sockets)x$($cpuInfo.P_Cores)" }
        $freqInfo = if ($cpuInfo.BaseFreq -and $cpuInfo.MaxFreq) { "$($cpuInfo.BaseFreq)/$($cpuInfo.MaxFreq)GHz" }
                    elseif ($cpuInfo.BaseFreq) { "$($cpuInfo.BaseFreq)GHz" }
                    elseif ($cpuInfo.MaxFreq) { "?/$($cpuInfo.MaxFreq)GHz" }
                    else { "" }
        Write-Output "$($cpuInfo.Name.Trim()) | $coreInfo $($cpuInfo.Threads) | $freqInfo | $($cpuInfo.SocketType)"

        $mem = $cpuInfo.MemoryInfo
        if ($mem.TotalSlots -gt 0) {
            Write-Output "$($mem.TotalGB) GB ($($mem.Config)) in $($mem.TotalSlots) slots, $($mem.UsedSlots) used"
        } else {
            Write-Output "$($mem.TotalGB) GB ($($mem.Config))"
        }
    }

    # Disks: physical (PS) + virtual (CPU-Z)
    $physDisks = Get-PhysicalDrivesList
    $virtDisks = Get-VirtualDrivesFromCpuZ -ReportContent $cpuzRaw
    $allDisks  = $physDisks + $virtDisks

    $i = 1
    $diskLines = $allDisks | ForEach-Object { "{0}. {1} {2} {3}" -f $i++, $_.MediaType, $_.Model, $_.Size }
    if ($diskLines) {
        Write-Output ($diskLines -join ' | ')
    }

    # Windows Version
    $winVer = Get-WindowsVersionString -ReportContent $cpuzRaw
    if ($winVer) {
        Write-Output "Windows Version: $winVer"
    }
} catch {
    Write-Error "Error: $_"
}
