# HwSummary.ps1
# Summary: Prints CPU, Memory, Disks, and Windows version
# Notes:
# - Single CPU‑Z run for all data
# - No blind sleeps: waits for process exit and for report file readiness
# - All user-visible text and comments are in English

# --- Install CPU-Z (winget/choco fallback) ---
function Install-CpuZ {
    $cpuZPath = "C:\Program Files\CPUID\CPU-Z\cpuz.exe"
    if (Test-Path $cpuZPath) { return $true }

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Error "Administrator privileges are required to install CPU-Z"
        return $false
    }

    try {
        Write-Host "Attempting to install CPU-Z via Winget..."
        winget install --id CPUID.CPU-Z -e --source winget -h
        if (Test-Path $cpuZPath) { return $true }
    } catch {
        Write-Warning "Winget failed or is unavailable. Falling back to Chocolatey..."
    }

    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Host "Installing Chocolatey..."
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Refresh-Environment
    }

    try {
        Write-Host "Installing CPU-Z via Chocolatey..."
        choco install cpu-z -y --force
        return $true
    } catch {
        Write-Error "Failed to install CPU-Z: $_"
        return $false
    }
}

# Refresh current environment (needed after choco install)
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

# --- CPU-Z report generation ---
# Single-run, waits for process exit and file readiness.
function Invoke-CpuZReport {
    if (-not (Install-CpuZ)) {
        throw "CPU-Z is not installed and cannot be installed automatically"
    }

    $cpuZPath   = "C:\Program Files\CPUID\CPU-Z\cpuz.exe"
    $reportBase = Join-Path $env:TEMP ("cpuz_report_" + ([guid]::NewGuid().Guid))
    $lowerPath  = "$reportBase.txt"
    $upperPath  = "$reportBase.TXT"

    Remove-Item $lowerPath,$upperPath -Force -ErrorAction SilentlyContinue

    Write-Host "Generating CPU-Z report..."
    $proc = Start-Process -FilePath $cpuZPath -ArgumentList "-txt=$reportBase" -WindowStyle Hidden -PassThru

    # Spinner while the process is alive
    $spinner = @('|','/','-','\'); $i = 0
    while (-not $proc.HasExited) {
        Write-Host -NoNewline ("`rGenerating CPU-Z report... " + $spinner[$i++ % $spinner.Length])
        Start-Sleep -Milliseconds 150
    }

    # Determine the path that was written
    $final = if (Test-Path $lowerPath) { $lowerPath } elseif (Test-Path $upperPath) { $upperPath } else {
        Write-Host ("`rGenerating CPU-Z report... failed".PadRight(60))
        throw "CPU-Z report file was not created"
    }

    Write-Host ("`rGenerating CPU-Z report... done".PadRight(60))

    # Wait until the file is readable and size is stable (handles brief locks)
    $stableTries = 0; $lastLen = -1; $start = Get-Date; $timeoutSec = 10
    while ($true) {
        try {
            $fi = Get-Item $final -ErrorAction Stop
            if ($fi.Length -gt 0 -and $fi.Length -eq $lastLen) { $stableTries++ } else { $stableTries = 0 }
            $lastLen = $fi.Length

            # non-locking open
            $fs = [System.IO.File]::Open($final, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $fs.Close()

            if ($stableTries -ge 2) { break }
        } catch {
            # still locked or not ready
        }
        if ((Get-Date) - $start -gt [timespan]::FromSeconds($timeoutSec)) {
            throw "CPU-Z report file is not readable after $timeoutSec seconds"
        }
        Start-Sleep -Milliseconds 150
    }

    # Robust read with a couple of retries
    $content = $null
    for ($t=0; $t -lt 5 -and -not $content; $t++) {
        try { $content = Get-Content $final -Raw -ErrorAction Stop } catch { Start-Sleep -Milliseconds 150 }
    }
    if (-not $content) { throw "Failed to read CPU-Z report" }

    Remove-Item $final -Force -ErrorAction SilentlyContinue
    return $content
}

# --- Parse CPU + Memory from CPU-Z report ---
function Get-CpuInfoFromCpuZ {
    param([Parameter(Mandatory=$true)][string]$ReportContent)

    $reportContent = $ReportContent

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

        # CPU name
        if ($cpuInfo -match "Name\s+(.+)") {
            $result.Name = $matches[1].Trim()
        }
        # Socket: "Socket 1700 LGA" -> "1700 LGA", "Socket AM5 (LGA1718)" -> "AM5"
        if ($cpuInfo -match 'Package(?:\s*\([^)]+\))?\s+Socket\s+([^(""\r\n]+)') {
            $result.SocketType = $matches[1].Trim()
        }
        # Hybrid Intel P/E cores
        if ($cpuInfo -match "Core Set 0\s+P-Cores, (\d+) cores") { $result.P_Cores = [int]$matches[1] }
        if ($cpuInfo -match "Core Set 1\s+E-Cores, (\d+) cores") { $result.E_Cores = [int]$matches[1] }
        # Non-hybrid (AMD/Xeon)
        if ($cpuInfo -match "Number of cores\s+(\d+)") {
            $cores = [int]$matches[1]
            if ($result.P_Cores -eq 0 -and $result.E_Cores -eq 0) { $result.P_Cores = $cores }
        }
        # Frequencies
        if ($cpuInfo -match "(?:Stock frequency|Base frequency \(cores\))\s+(\d+) MHz") {
            $baseFreq = [decimal]$matches[1] / 1000
            $result.BaseFreq = $baseFreq.ToString("0.0").Replace(".", ",")
        }
        if ($cpuInfo -match "Max frequency\s+(\d+) MHz") {
            $maxFreq = [decimal]$matches[1] / 1000
            $result.MaxFreq = $maxFreq.ToString("0.0").Replace(".", ",")
        }
    }

    # --- Memory parsing ---
    $memoryInfo = [PSCustomObject]@{
        TotalGB     = 0
        Config      = ""
        TotalSlots  = 0
        UsedSlots   = 0
        RawModules  = @()
    }

    # Match all "DMI Memory Device" blocks
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

    # Sum all "max# of devices" across arrays (systems with multiple memory arrays)
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

# --- Parse Virtual Drives from CPU-Z report ---
function Get-VirtualDrivesFromCpuZ {
    param([Parameter(Mandatory=$true)][string]$ReportContent)

    $virtual = @()
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

# --- Physical disks from Windows (excluding obvious virtuals) ---
function Get-PhysicalDrivesList {
    $list = @()
    try { $pd = Get-PhysicalDisk } catch { return $list }

    foreach ($d in $pd) {
        $model = if ($d.PSObject.Properties.Name -contains 'Model' -and $d.Model) { $d.Model }
                 elseif ($d.FriendlyName) { $d.FriendlyName }
                 else { '' }

        if ($model -match '(?i)\b(QEMU|VBOX|VMWARE|VIRTIO|HYPER-V|KVM|MSFT VIRTUAL|VIRTUAL)\b') { continue }

        $type  = if ($d.MediaType) { $d.MediaType } else { 'Unspecified' }
        $bytes = $d.Size
        $sizeStr = if ($bytes -ge 1e12) { "{0:N1} TB" -f ($bytes/1e12) } else { "{0:N1} GB" -f ($bytes/1e9) }
        $list += [pscustomobject]@{ IsVirtual = $false; MediaType = $type; Model = $model; Size = $sizeStr }
    }
    return $list
}

# --- Windows Version ---
function Get-WindowsVersionString {
    param([string]$ReportContent)
    if ($ReportContent -and ($ReportContent -match '(?mi)^\s*Windows\s+Version\s+(.+)$')) {
        return $matches[1].Trim()
    }
    try {
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $product = $cv.ProductName
        $display = $cv.DisplayVersion; if (-not $display) { $display = $cv.ReleaseId }
        $build   = $cv.CurrentBuild; $ubr = $cv.UBR
        $buildStr = if ($ubr -ge 0) { "$build.$ubr" } else { "$build" }
        return "$product $display ($buildStr)"
    } catch { return "" }
}

# --- MAIN ---
try {
    $cpuzRaw = Invoke-CpuZReport
    if ([string]::IsNullOrWhiteSpace($cpuzRaw)) { throw "CPU-Z report content is empty" }

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

    # Disks: physical (PowerShell) + virtual (CPU-Z)
    $physDisks = Get-PhysicalDrivesList
    $virtDisks = Get-VirtualDrivesFromCpuZ -ReportContent $cpuzRaw
    $allDisks  = $physDisks + $virtDisks
    $i = 1
    $diskLines = $allDisks | ForEach-Object { "{0}. {1} {2} {3}" -f $i++, $_.MediaType, $_.Model, $_.Size }
    if ($diskLines) { Write-Output ($diskLines -join ' | ') }

    # Windows Version
    $winVer = Get-WindowsVersionString -ReportContent $cpuzRaw
    if ($winVer) { Write-Output "Windows Version: $winVer" }

} catch {
    Write-Error "Error: $_"
}
