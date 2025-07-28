# HwSummary.ps1
# Summary: Displays CPU, Memory, Disk, and Windows version info
# Requires: CPU-Z (auto-installs if missing)

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
function Invoke-CpuZReport {
    if (-not (Install-CpuZ)) {
        throw "CPU-Z is not installed and cannot be installed automatically"
    }

    $cpuZPath = "C:\Program Files\CPUID\CPU-Z\cpuz.exe"
    $reportBase = Join-Path $env:TEMP ("cpuz_report_" + ([guid]::NewGuid().Guid))
    $lowerPath = "$reportBase.txt"
    $upperPath = "$reportBase.TXT"
    Remove-Item $lowerPath,$upperPath -Force -ErrorAction SilentlyContinue

    Write-Host "Generating CPU-Z report..."
    $proc = Start-Process -FilePath $cpuZPath -ArgumentList "-txt=$reportBase" -WindowStyle Hidden -PassThru

    $spinner = @('|','/','-','\'); $i = 0
    while (-not $proc.HasExited) {
        Write-Host -NoNewline ("`rGenerating CPU-Z report... " + $spinner[$i++ % $spinner.Length])
        Start-Sleep -Milliseconds 150
    }

    $final = if (Test-Path $lowerPath) { $lowerPath } elseif (Test-Path $upperPath) { $upperPath } else {
        Write-Host ("`rGenerating CPU-Z report... failed".PadRight(60))
        throw "CPU-Z report file was not created"
    }

    Write-Host ("`rGenerating CPU-Z report... done".PadRight(60))
    $stableTries = 0; $lastLen = -1; $start = Get-Date; $timeoutSec = 10
    while ($true) {
        try {
            $fi = Get-Item $final -ErrorAction Stop
            if ($fi.Length -gt 0 -and $fi.Length -eq $lastLen) { $stableTries++ } else { $stableTries = 0 }
            $lastLen = $fi.Length
            $fs = [System.IO.File]::Open($final, 'Open', 'Read', 'ReadWrite'); $fs.Close()
            if ($stableTries -ge 2) { break }
        } catch {}
        if ((Get-Date) - $start -gt [timespan]::FromSeconds($timeoutSec)) {
            throw "CPU-Z report file is not readable after $timeoutSec seconds"
        }
        Start-Sleep -Milliseconds 150
    }

    $content = $null
    for ($t=0; $t -lt 5 -and -not $content; $t++) {
        try { $content = Get-Content $final -Raw -ErrorAction Stop } catch { Start-Sleep -Milliseconds 150 }
    }
    if (-not $content) { throw "Failed to read CPU-Z report" }
    Remove-Item $final -Force -ErrorAction SilentlyContinue
    return $content
}

# --- Helpers to extract CPU, Memory, Virtual drives, etc ---
function Get-CpuInfoFromCpuZ { param([string]$ReportContent) if (-not $ReportContent) { return $null }
# Same code block you already have from previous message — omitted here for brevity
}

function Get-VirtualDrivesFromCpuZ { param([string]$ReportContent) if (-not $ReportContent) { return @() }
# Same block as before — omitted for brevity
}

function Get-PhysicalDrivesList {
    $list = @(); try { $pd = Get-PhysicalDisk } catch { return $list }
    foreach ($d in $pd) {
        $model = if ($d.PSObject.Properties.Name -contains 'Model') { $d.Model } elseif ($d.FriendlyName) { $d.FriendlyName } else { '' }
        if ($model -match '(?i)QEMU|VBOX|VMWARE|VIRTUAL|HYPER-V|KVM') { continue }
        $type = if ($d.MediaType) { $d.MediaType } else { 'Unspecified' }
        $size = if ($d.Size -ge 1e12) { "{0:N1} TB" -f ($d.Size/1e12) } else { "{0:N1} GB" -f ($d.Size/1e9) }
        $list += [pscustomobject]@{ IsVirtual = $false; MediaType = $type; Model = $model; Size = $size }
    }
    return $list
}

function Get-WindowsVersionString {
    param([string]$ReportContent)
    if ($ReportContent -match '(?mi)^\s*Windows\s+Version\s+(.+)$') { return $matches[1].Trim() }
    try {
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $product = $cv.ProductName
        $display = $cv.DisplayVersion; if (-not $display) { $display = $cv.ReleaseId }
        $build = $cv.CurrentBuild; $ubr = $cv.UBR
        return "$product $display ($build.$ubr)"
    } catch { return "" }
}

# --- MAIN ---
try {
    $cpuzRaw = Invoke-CpuZReport
    if ([string]::IsNullOrWhiteSpace($cpuzRaw)) { throw "CPU-Z report content is empty" }

    $cpuInfo = Get-CpuInfoFromCpuZ -ReportContent $cpuzRaw
    if ($cpuInfo) {
        $coreInfo = if ($cpuInfo.E_Cores -gt 0) { "$($cpuInfo.Sockets)x$($cpuInfo.P_Cores)P+$($cpuInfo.E_Cores)E" } else { "$($cpuInfo.Sockets)x$($cpuInfo.P_Cores)" }
        $freqInfo = if ($cpuInfo.BaseFreq -and $cpuInfo.MaxFreq) { "$($cpuInfo.BaseFreq)/$($cpuInfo.MaxFreq)GHz" } elseif ($cpuInfo.BaseFreq) { "$($cpuInfo.BaseFreq)GHz" } elseif ($cpuInfo.MaxFreq) { "?/$($cpuInfo.MaxFreq)GHz" } else { "" }
        Write-Output "$($cpuInfo.Name.Trim()) | $coreInfo $($cpuInfo.Threads) | $freqInfo | $($cpuInfo.SocketType)"
        $mem = $cpuInfo.MemoryInfo
        if ($mem.TotalSlots -gt 0) {
            Write-Output "$($mem.TotalGB) GB ($($mem.Config)) in $($mem.TotalSlots) slots, $($mem.UsedSlots) used"
        } else {
            Write-Output "$($mem.TotalGB) GB ($($mem.Config))"
        }
    }

    $physDisks = Get-PhysicalDrivesList
    $virtDisks = Get-VirtualDrivesFromCpuZ -ReportContent $cpuzRaw
    $allDisks  = $physDisks + $virtDisks
    $i = 1
    $diskLines = $allDisks | ForEach-Object { "{0}. {1} {2} {3}" -f $i++, $_.MediaType, $_.Model, $_.Size }
    if ($diskLines) { Write-Output ($diskLines -join ' | ') }

    $winVer = Get-WindowsVersionString -ReportContent $cpuzRaw
    if ($winVer) { Write-Output "Windows Version: $winVer" }

} catch {
    Write-Error "Error: $_"
}
