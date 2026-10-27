[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Command = "help"
)

$Version = "0.7.0"

function Get-DcInventory {
    $observedAt = Get-Date

    try {
        $pnpDevices = @(Get-PnpDevice -ErrorAction Stop)
    } catch {
        throw "Не удалось прочитать список устройств (Get-PnpDevice): $($_.Exception.Message)"
    }

    try {
        $signedDrivers = @(Get-CimInstance Win32_PnPSignedDriver -ErrorAction Stop)
    } catch {
        throw "Не удалось прочитать сведения о драйверах (Win32_PnPSignedDriver): $($_.Exception.Message)"
    }

    $driversById = @{}
    foreach ($driver in $signedDrivers) {
        if ([string]::IsNullOrWhiteSpace($driver.DeviceID)) { continue }
        $key = $driver.DeviceID.ToUpperInvariant()
        if (-not $driversById.ContainsKey($key)) {
            $driversById[$key] = @()
        }
        $driversById[$key] = @($driversById[$key]) + $driver
    }

    foreach ($device in $pnpDevices) {
        $instanceId = [string]$device.InstanceId
        if ([string]::IsNullOrWhiteSpace($instanceId)) { continue }

        $driverMatches = @()
        $key = $instanceId.ToUpperInvariant()
        if ($driversById.ContainsKey($key)) {
            $driverMatches = @($driversById[$key])
        }

        $driver = $null
        $dataQuality = "Partial"
        if ($driverMatches.Count -eq 1) {
            $driver = $driverMatches[0]
            $dataQuality = "Ok"
            if ([string]::IsNullOrWhiteSpace($driver.DriverVersion) -or [string]::IsNullOrWhiteSpace($driver.InfName)) {
                $dataQuality = "Partial"
            }
        } elseif ($driverMatches.Count -gt 1) {
            $driver = $driverMatches[0]
            $dataQuality = "Conflict"
        }

        $problemCode = $null
        if ($null -ne $device.ConfigManagerErrorCode) {
            $problemCode = [int]$device.ConfigManagerErrorCode
        }

        $status = "Unknown"
        $problemDescription = "Состояние устройства неизвестно"
        if ($null -ne $problemCode) {
            switch ($problemCode) {
                0 {
                    $status = "OK"
                    $problemDescription = "Устройство работает нормально"
                }
                45 {
                    $status = "NotPresent"
                    $problemDescription = "Устройство не подключено (код 45)"
                }
                22 {
                    $status = "Disabled"
                    $problemDescription = "Устройство отключено (код 22)"
                }
                default {
                    $status = "Problem"
                    $problemDescription = "Проблема устройства (код $problemCode)"
                    if (-not [string]::IsNullOrWhiteSpace($device.ProblemDescription)) {
                        $problemDescription = "$($device.ProblemDescription) (код $problemCode)"
                    }
                }
            }
        }

        $presence = "Unknown"
        if ($null -ne $device.Present) {
            if ($device.Present -eq $true) {
                $presence = "Present"
            } else {
                $presence = "NotPresent"
            }
        } elseif ($problemCode -eq 45) {
            $presence = "NotPresent"
        } elseif ($null -ne $problemCode) {
            $presence = "Present"
        }

        $friendlyName = $instanceId
        if (-not [string]::IsNullOrWhiteSpace($device.FriendlyName)) {
            $friendlyName = $device.FriendlyName
        }

        $deviceClass = "Unknown"
        if (-not [string]::IsNullOrWhiteSpace($device.Class)) {
            $deviceClass = $device.Class
        }

        $deviceId = $device.DeviceID
        $manufacturer = $null
        $driverVersion = $null
        $driverDate = $null
        $infName = $null
        $isSigned = $null
        $signer = $null
        if ($driver) {
            $deviceId = $driver.DeviceID
            $manufacturer = $driver.Manufacturer
            if ([string]::IsNullOrWhiteSpace($manufacturer)) {
                $manufacturer = $driver.DriverProviderName
            }
            $driverVersion = $driver.DriverVersion
            $driverDate = $driver.DriverDate
            $infName = $driver.InfName
            $isSigned = $driver.IsSigned
            $signer = $driver.Signer
        }

        [PSCustomObject]@{
            InstanceId = $instanceId
            DeviceId = $deviceId
            Class = $deviceClass
            FriendlyName = $friendlyName
            Manufacturer = $manufacturer
            Presence = $presence
            ProblemCode = $problemCode
            Status = $status
            ProblemDescription = $problemDescription
            DriverVersion = $driverVersion
            DriverDate = $driverDate
            InfName = $infName
            IsSigned = $isSigned
            Signer = $signer
            DataQuality = $dataQuality
            ObservedAt = $observedAt
        }
    }
}

function Format-DcTable {
    param([object[]]$Devices)

    $Devices |
        Select-Object Class, FriendlyName,
            @{ Name = "Manufacturer"; Expression = {
                if ($_.Manufacturer) { $_.Manufacturer } else { "—" }
            } },
            @{ Name = "DriverVersion"; Expression = {
                if ($_.DriverVersion) { $_.DriverVersion } else { "—" }
            } },
            @{ Name = "DriverDate"; Expression = {
                if ($_.DriverDate) {
                    ([datetime]$_.DriverDate).ToString("yyyy-MM-dd")
                } else { "—" }
            } },
            @{ Name = "InfName"; Expression = {
                if ($_.InfName) { $_.InfName } else { "—" }
            } },
            DataQuality |
        Format-Table -AutoSize | Out-String -Width 220
}

function Show-DriverList {
    $devices = @(Get-DcInventory |
        Where-Object { $_.Presence -ne "NotPresent" } |
        Sort-Object Class, FriendlyName)
    Format-DcTable $devices
    Write-Host "Всего записей: $($devices.Count)"
    Write-Host "Не подключённые сейчас устройства (код 45) в список не входят."
}

function Show-DeviceScan {
    $devices = @(Get-DcInventory)

    $ok = @($devices | Where-Object Status -eq "OK")
    $notPresent = @($devices | Where-Object Status -eq "NotPresent")
    $disabled = @($devices | Where-Object Status -eq "Disabled")
    $problem = @($devices | Where-Object Status -eq "Problem")
    $unknown = @($devices | Where-Object Status -eq "Unknown")
    $partial = @($devices | Where-Object DataQuality -eq "Partial")
    $conflict = @($devices | Where-Object DataQuality -eq "Conflict")
    $partialPresent = @($partial | Where-Object Presence -eq "Present")
    $partialNotPresent = @($partial | Where-Object Presence -eq "NotPresent")
    $partialUnknownPresence = @($partial | Where-Object Presence -eq "Unknown")
    $script:DcProblemCount = $problem.Count

    Write-Host "Устройств проверено: $($devices.Count)"
    Write-Host "Работают без ошибок (код 0): $($ok.Count)"
    Write-Host "Не подключены сейчас (код 45): $($notPresent.Count) — это не ошибка"
    Write-Host "Отключены пользователем (код 22): $($disabled.Count)"
    Write-Host "Требуют внимания: $($problem.Count)"
    Write-Host "Состояние неизвестно: $($unknown.Count)"

    if ($problem.Count -gt 0) {
        $problem |
            Select-Object Class, FriendlyName, ProblemCode, ProblemDescription |
            Format-Table -AutoSize | Out-String -Width 220
        Write-Host "Частые коды: 28 = драйвер не установлен, 43 = ошибка устройства."
    } else {
        Write-Host "Устройств с ошибками драйверов не найдено."
    }

    if ($unknown.Count -gt 0) {
        Write-Host "Устройства с неизвестным состоянием (это не доказанная ошибка):"
        $unknown |
            Select-Object Class, FriendlyName, ProblemDescription |
            Format-Table -AutoSize | Out-String -Width 220
    }

    Write-Host "Неполные данные о драйвере: $($partial.Count) — это не ошибка устройства"
    Write-Host "  из них присутствуют (Present): $($partialPresent.Count)"
    Write-Host "  из них не подключены (NotPresent): $($partialNotPresent.Count)"
    if ($partialUnknownPresence.Count -gt 0) {
        Write-Host "  из них присутствие неизвестно: $($partialUnknownPresence.Count)"
    }
    Write-Host "Конфликтующие записи о драйвере: $($conflict.Count)"
    if ($partial.Count -gt 0 -or $conflict.Count -gt 0) {
        @($partial + $conflict) |
            Select-Object Class, FriendlyName, Presence, Status, DataQuality |
            Format-Table -AutoSize | Out-String -Width 220
    }
}

function Show-SystemInfo {
    $cs = Get-CimInstance Win32_ComputerSystem
    $os = Get-CimInstance Win32_OperatingSystem
    $cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
    $devices = @(Get-DcInventory)

    $present = @($devices | Where-Object Presence -eq "Present")
    $notPresent = @($devices | Where-Object Status -eq "NotPresent")
    $unknownPresence = @($devices | Where-Object Presence -eq "Unknown")
    $partial = @($devices | Where-Object DataQuality -eq "Partial")
    $conflict = @($devices | Where-Object DataQuality -eq "Conflict")

    $observedAtText = "неизвестно"
    if ($devices.Count -gt 0) {
        $observedAtText = $devices[0].ObservedAt.ToString("yyyy-MM-dd HH:mm:ss")
    }

    [PSCustomObject]@{
        "Производитель" = $cs.Manufacturer
        "Модель" = $cs.Model
        "Windows" = $os.Caption
        "Сборка ОС" = $os.BuildNumber
        "Разрядность" = $os.OSArchitecture
        "Процессор" = $cpu.Name
        "Ядра / потоки" = "$($cpu.NumberOfCores) / $($cpu.NumberOfLogicalProcessors)"
        "Память, ГБ" = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
        "Устройств в инвентаре" = $devices.Count
        "Присутствуют сейчас" = $present.Count
        "Не подключены (код 45)" = $notPresent.Count
        "Присутствие неизвестно" = $unknownPresence.Count
        "Неполные данные о драйвере" = $partial.Count
        "Конфликтующие записи о драйвере" = $conflict.Count
        "Источники данных" = "Get-PnpDevice, Win32_PnPSignedDriver"
        "Время сбора" = $observedAtText
        "Режим" = "только чтение"
    } | Format-List
}

function Show-Help {
    @"
DriverCheck $Version — диагностика драйверов Windows (только чтение).

Использование:
  drivercheck.ps1 <команда>

Команды:
  help      Показать эту справку
  list      Устройства и установленные драйверы
  scan      Проверка устройств на ошибки
  info      Сведения о системе

Коды возврата:
  0  успешно; для scan — проблем не найдено
  1  scan нашёл устройства, требующие внимания
  2  неизвестная команда
  3  не удалось получить данные об устройствах

Проверка обновлений пока не реализована: нужен надёжный источник сравнения.
Программа не устанавливает, не удаляет и не изменяет драйверы.
"@
}

$exitCode = 0

try {
    switch ($Command.ToLowerInvariant()) {
        "help" { Show-Help }
        "list" { Show-DriverList }
        "scan" {
            Show-DeviceScan
            if ($script:DcProblemCount -gt 0) {
                $exitCode = 1
            }
        }
        "info" { Show-SystemInfo }
        default {
            Write-Host "Неизвестная команда: '$Command'."
            Show-Help
            $exitCode = 2
        }
    }
} catch {
    Write-Host "Ошибка получения данных: $($_.Exception.Message)"
    Write-Host "Команда не выполнена. Код возврата: 3."
    $exitCode = 3
}

exit $exitCode
