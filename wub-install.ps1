# Скрипт для загрузки Windows Update Blocker и добавления в PATH

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
    Write-Host "[0] Пропустить" -ForegroundColor Gray
    Write-Host ""
}

function Get-UserChoice {
    param([int]$MaxOption)
    
    do {
        $choice = Read-Host "Выберите опцию (0-$MaxOption)"
        if ($choice -match '^\d+$' -and [int]$choice -ge 0 -and [int]$choice -le $MaxOption) {
            return [int]$choice
        }
        Write-Host "Неверный выбор! Введите число от 0 до $MaxOption" -ForegroundColor Red
    } while ($true)
}

function Download-File {
    param([string]$Url, [string]$FilePath)
    
    try {
        Write-Status "Загрузка файла..."
        Write-Status "URL: $Url" "Cyan"
        Write-Status "Назначение: $FilePath" "Cyan"
        
        # Создание WebClient с настройками
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        
        # Загрузка файла
        $webClient.DownloadFile($Url, $FilePath)
        $webClient.Dispose()
        
        # Проверка успешности загрузки
        if (Test-Path $FilePath) {
            $fileSize = (Get-Item $FilePath).Length
            Write-Status "Файл успешно загружен! Размер: $([Math]::Round($fileSize/1KB, 2)) KB"
            return $true
        } else {
            Write-Status "Ошибка: файл не найден после загрузки" "Red"
            return $false
        }
    }
    catch {
        Write-Status "Ошибка загрузки: $($_.Exception.Message)" "Red"
        
        # Альтернативный метод через Invoke-WebRequest
        try {
            Write-Status "Попытка альтернативного метода загрузки..." "Yellow"
            Invoke-WebRequest -Uri $Url -OutFile $FilePath -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            
            if (Test-Path $FilePath) {
                Write-Status "Файл загружен альтернативным методом!"
                return $true
            }
        }
        catch {
            Write-Status "Альтернативный метод также не сработал: $($_.Exception.Message)" "Red"
        }
        
        return $false
    }
}

# Основная логика
Write-Status "=== Загрузчик Windows Update Blocker ===" "Cyan"

# Создание целевой папки если её нет
$fullFilePath = Join-Path $DestinationPath $FileName

try {
    if (-not (Test-Path $DestinationPath)) {
        Write-Status "Создание папки: $DestinationPath"
        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    }
}
catch {
    Write-Status "Ошибка создания папки: $($_.Exception.Message)" "Red"
    exit 1
}

# Проверка существования файла
if (Test-Path $fullFilePath) {
    Write-Status "Файл уже существует: $fullFilePath" "Yellow"
    
    $fileInfo = Get-Item $fullFilePath
    Write-Status "Размер файла: $([Math]::Round($fileInfo.Length/1KB, 2)) KB"
    Write-Status "Дата создания: $($fileInfo.CreationTime)"
    
    $overwrite = Read-Host "Перезаписать файл? (y/N)"
    if ($overwrite -notmatch '^[yY]') {
        Write-Status "Загрузка пропущена"
        $downloadSuccess = $true
    } else {
        $downloadSuccess = Download-File -Url $DownloadUrl -FilePath $fullFilePath
    }
} else {
    Write-Status "Файл не найден, начинаем загрузку..."
    $downloadSuccess = Download-File -Url $DownloadUrl -FilePath $fullFilePath
}

if (-not $downloadSuccess) {
    Write-Status "Не удалось загрузить файл. Завершение работы." "Red"
    exit 1
}

# Проверка возможности добавления в PATH
Write-Status ""
Write-Status "=== Настройка переменной PATH ===" "Cyan"

# Проверка текущего PATH
$currentPath = $env:PATH
$targetPath = $DestinationPath

if ($currentPath -like "*$targetPath*") {
    Write-Status "Путь уже добавлен в переменную PATH" "Yellow"
} else {
    Write-Status "Путь не найден в переменной PATH"
    
    # Проверка прав администратора для системной переменной
    $isAdmin = Test-AdminRights
    if ($isAdmin) {
        Write-Status "Обнаружены права администратора" "Green"
    } else {
        Write-Status "Права администратора не обнаружены" "Yellow"
        Write-Status "Для изменения системной переменной PATH требуются права администратора" "Yellow"
    }
    
    # Меню выбора действий
    $menuOptions = @()
    $commands = @()
    
    if ($isAdmin) {
        $menuOptions += "Добавить в системную переменную PATH (для всех пользователей)"
        $commands += '[Environment]::SetEnvironmentVariable("PATH", $env:PATH + ";' + $targetPath + '", [EnvironmentVariableTarget]::Machine)'
    }
    
    $menuOptions += "Добавить в пользовательскую переменную PATH (только текущий пользователь)"
    $commands += '[Environment]::SetEnvironmentVariable("PATH", $env:PATH + ";' + $targetPath + '", [EnvironmentVariableTarget]::User)'
    
    $menuOptions += "Добавить только в текущую сессию PowerShell"
    $commands += '$env:PATH += ";' + $targetPath + '"'
    
    $menuOptions += "Показать команды для ручного выполнения"
    
    Show-Menu -Title "Выберите способ добавления в PATH" -Options $menuOptions
    $choice = Get-UserChoice -MaxOption $menuOptions.Count
    
    if ($choice -eq 0) {
        Write-Status "Настройка PATH пропущена"
    } elseif ($choice -eq $menuOptions.Count) {
        # Показать команды
        Write-Status ""
        Write-Status "=== Команды для ручного выполнения ===" "Cyan"
        Write-Status ""
        
        if ($isAdmin) {
            Write-Status "1. Для системной переменной PATH (требуются права администратора):" "Yellow"
            Write-Host '[Environment]::SetEnvironmentVariable("PATH", $env:PATH + ";C:\Users\admin\Desktop\Wub", [EnvironmentVariableTarget]::Machine)' -ForegroundColor White
            Write-Status ""
        }
        
        Write-Status "2. Для пользовательской переменной PATH:" "Yellow"
        Write-Host '[Environment]::SetEnvironmentVariable("PATH", $env:PATH + ";C:\Users\admin\Desktop\Wub", [EnvironmentVariableTarget]::User)' -ForegroundColor White
        Write-Status ""
        
        Write-Status "3. Для текущей сессии PowerShell:" "Yellow"
        Write-Host '$env:PATH += ";C:\Users\admin\Desktop\Wub"' -ForegroundColor White
        Write-Status ""
        
    } else {
        # Выполнить выбранную команду
        $selectedCommand = $commands[$choice - 1]
        
        Write-Status "Выполнение команды:" "Yellow"
        Write-Host $selectedCommand -ForegroundColor White
        
        $confirm = Read-Host "Подтвердите выполнение (Y/n)"
        if ($confirm -notmatch '^[nN]') {
            try {
                Invoke-Expression $selectedCommand
                Write-Status "Команда выполнена успешно!" "Green"
                
                # Если это была системная переменная, добавляем также в текущую сессию
                if ($choice -eq 1 -and $isAdmin) {
                    Write-Status "Добавление пути в текущую сессию..."
                    $env:PATH += ";$targetPath"
                }
                
                Write-Status "Проверка доступности команды Wub..."
                try {
                    $wubVersion = & "$fullFilePath" 2>$null
                    Write-Status "Команда Wub доступна!" "Green"
                }
                catch {
                    Write-Status "Команда Wub пока недоступна. Возможно, потребуется перезапуск PowerShell." "Yellow"
                }
                
            }
            catch {
                Write-Status "Ошибка выполнения команды: $($_.Exception.Message)" "Red"
            }
        } else {
            Write-Status "Выполнение команды отменено"
        }
    }
}

Write-Status ""
Write-Status "=== Завершение работы ===" "Cyan"
Write-Status "Файл Windows Update Blocker находится здесь: $fullFilePath"

if (Test-Path $fullFilePath) {
    Write-Status "Размер файла: $([Math]::Round((Get-Item $fullFilePath).Length/1KB, 2)) KB"
}

Write-Status ""
Write-Status "Для использования из любого места выполните одну из команд:" "Yellow"
Write-Status "wub /D  - отключить обновления" "Cyan"
Write-Status "wub /E  - включить обновления" "Cyan"
Write-Status ""

Read-Host "Нажмите Enter для завершения"
