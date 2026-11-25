@echo off
setlocal EnableExtensions
title Kiem tra he thong va driver - Windows 11

:: ===== Yeu cau quyen Administrator =====
net session >nul 2>&1
if %errorlevel% neq 0 goto :elevate
goto :main

:elevate
echo Dang yeu cau quyen Administrator...
powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
exit /b

:main
cls
echo ============================================================
echo    KIEM TRA TINH TRANG HE THONG VA DRIVER - WINDOWS 11
echo ============================================================

:: Duong dan bao cao tren Desktop
for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"`) do set "DESK=%%i"
for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"`) do set "TS=%%i"
set "REPORT=%DESK%\SystemCheck_%COMPUTERNAME%_%TS%.txt"
powershell -NoProfile -Command "Set-Content -LiteralPath $env:REPORT -Value ('SYSTEM CHECK REPORT - ' + (Get-Date)) -Encoding UTF8"

:: ===== 1. He dieu hanh va phan cung =====
call :ps "1. HE DIEU HANH VA PHAN CUNG" "Get-CimInstance Win32_OperatingSystem | Select-Object Caption,Version,BuildNumber,OSArchitecture,InstallDate,LastBootUpTime | Format-List; Get-CimInstance Win32_ComputerSystem | Select-Object Manufacturer,Model,@{n='RAM_GB';e={[math]::Round($_.TotalPhysicalMemory/1GB,1)}} | Format-List; Get-CimInstance Win32_BIOS | Select-Object SMBIOSBIOSVersion,ReleaseDate | Format-List"

:: ===== 2. CPU, RAM, thoi gian chay =====
call :ps "2. CPU - RAM - UPTIME" "Get-CimInstance Win32_Processor | Select-Object Name,NumberOfCores,NumberOfLogicalProcessors,LoadPercentage | Format-List; Get-CimInstance Win32_OperatingSystem | Select-Object @{n='RAM_Total_GB';e={[math]::Round($_.TotalVisibleMemorySize/1MB,1)}},@{n='RAM_Free_GB';e={[math]::Round($_.FreePhysicalMemory/1MB,1)}} | Format-List; $u=(Get-Date)-(Get-CimInstance Win32_OperatingSystem).LastBootUpTime; ('Uptime: {0} ngay {1} gio {2} phut' -f $u.Days,$u.Hours,$u.Minutes)"

:: ===== 3. O dia =====
call :ps "3. O DIA (SUC KHOE VA DUNG LUONG)" "Get-PhysicalDisk | Select-Object FriendlyName,MediaType,HealthStatus,OperationalStatus,@{n='Size_GB';e={[math]::Round($_.Size/1GB,0)}} | Format-Table -AutoSize; Get-PhysicalDisk | Get-StorageReliabilityCounter | Select-Object DeviceId,Temperature,Wear,PowerOnHours,ReadErrorsTotal,WriteErrorsTotal | Format-Table -AutoSize; Get-Volume | Where-Object DriveLetter | Select-Object DriveLetter,FileSystemLabel,FileSystem,HealthStatus,@{n='Size_GB';e={[math]::Round($_.Size/1GB,1)}},@{n='Free_GB';e={[math]::Round($_.SizeRemaining/1GB,1)}} | Format-Table -AutoSize"

:: ===== 4. Windows Update =====
call :ps "4. WINDOWS UPDATE" "Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 10 HotFixID,Description,InstalledOn | Format-Table -AutoSize; [pscustomobject]@{WU_RebootRequired=(Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'); CBS_RebootPending=(Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending')} | Format-List"

:: ===== 5. Dich vu he thong quan trong =====
call :ps "5. DICH VU QUAN TRONG" "Get-Service wuauserv,bits,WinDefend,EventLog,Winmgmt,mpssvc,Dhcp,Dnscache,Audiosrv | Select-Object Name,Status,StartType | Format-Table -AutoSize"

:: ===== 6. Bao mat =====
call :ps "6. WINDOWS DEFENDER VA FIREWALL" "Get-MpComputerStatus | Select-Object AMServiceEnabled,AntivirusEnabled,RealTimeProtectionEnabled,AntivirusSignatureLastUpdated,QuickScanEndTime | Format-List; Get-NetFirewallProfile | Select-Object Name,Enabled | Format-Table -AutoSize"

:: ===== 7. Thiet bi loi driver =====
call :ps "7. THIET BI BI LOI DRIVER" "$p = Get-PnpDevice -PresentOnly | Where-Object { $_.Status -ne 'OK' }; if ($p) { $p | Select-Object Status,Class,FriendlyName,InstanceId,@{n='ErrorCode';e={(Get-PnpDeviceProperty -InstanceId $_.InstanceId -KeyName 'DEVPKEY_Device_ProblemCode').Data}} | Format-Table -AutoSize -Wrap } else { 'OK: Khong co thiet bi nao bi loi driver.' }"

:: ===== 8. Driver chinh =====
call :ps "8. DRIVER CHINH (VGA, MANG, AM THANH, BLUETOOTH, O DIA...)" "Get-CimInstance Win32_PnPSignedDriver | Where-Object { $_.DeviceName -and $_.DeviceClass -in 'DISPLAY','NET','MEDIA','BLUETOOTH','SCSIADAPTER','HDC','CAMERA','FIRMWARE' } | Sort-Object DeviceClass,DeviceName | Select-Object DeviceClass,DeviceName,Manufacturer,DriverVersion,@{n='DriverDate';e={ if ($_.DriverDate) { $_.DriverDate.ToString('yyyy-MM-dd') } }} | Format-Table -AutoSize -Wrap"

:: ===== 9. Driver cu / khong co chu ky =====
call :ps "9. DRIVER HANG THU BA CU HON 3 NAM" "$d = Get-CimInstance Win32_PnPSignedDriver | Where-Object { $_.DeviceName -and $_.DriverDate -and $_.DriverDate -lt (Get-Date).AddYears(-3) -and $_.Manufacturer -notlike 'Microsoft*' } | Sort-Object DriverDate; if ($d) { $d | Select-Object DeviceClass,DeviceName,Manufacturer,DriverVersion,@{n='DriverDate';e={$_.DriverDate.ToString('yyyy-MM-dd')}} | Format-Table -AutoSize -Wrap } else { 'OK: Khong co driver nao cu hon 3 nam.' }"

call :ps "10. DRIVER KHONG CO CHU KY SO" "$u = Get-CimInstance Win32_PnPSignedDriver | Where-Object { $_.DeviceName -and $_.IsSigned -eq $false }; if ($u) { $u | Select-Object DeviceClass,DeviceName,Manufacturer,DriverVersion | Format-Table -AutoSize -Wrap } else { 'OK: Tat ca driver deu co chu ky.' }"

:: ===== 11. Card do hoa =====
call :ps "11. CARD DO HOA" "Get-CimInstance Win32_VideoController | Select-Object Name,DriverVersion,@{n='DriverDate';e={ if ($_.DriverDate) { $_.DriverDate.ToString('yyyy-MM-dd') } }},Status,CurrentHorizontalResolution,CurrentVerticalResolution | Format-List"

:: ===== 12. Mang =====
call :ps "12. MANG" "Get-NetAdapter | Where-Object Status -eq 'Up' | Select-Object Name,InterfaceDescription,LinkSpeed | Format-Table -AutoSize; [pscustomobject]@{Ping_8_8_8_8=(Test-Connection -ComputerName 8.8.8.8 -Count 2 -Quiet); Ping_google_com=(Test-Connection -ComputerName google.com -Count 2 -Quiet)} | Format-List"

:: ===== 13. Nhat ky loi 7 ngay gan nhat =====
call :ps "13. LOI NGHIEM TRONG TRONG SYSTEM LOG (7 NGAY)" "$e = Get-WinEvent -FilterHashtable @{LogName='System';Level=1,2;StartTime=(Get-Date).AddDays(-7)} -MaxEvents 20; if ($e) { $e | Select-Object TimeCreated,ProviderName,Id,@{n='Message';e={ ($_.Message -split '\r?\n')[0] }} | Format-Table -AutoSize -Wrap } else { 'OK: Khong co loi Critical/Error trong 7 ngay.' }"

call :ps "14. LAN MAN HINH XANH (BUGCHECK) GAN DAY" "$b = Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-WER-SystemErrorReporting';Id=1001} -MaxEvents 5; if ($b) { $b | Select-Object TimeCreated,@{n='Message';e={ ($_.Message -split '\r?\n')[0] }} | Format-Table -AutoSize -Wrap } else { 'OK: Khong co BSOD gan day.' }"

:: ===== 15. Pin (neu la laptop) =====
powershell -NoProfile -Command "if (Get-CimInstance Win32_Battery) { exit 0 } else { exit 1 }"
if %errorlevel%==0 (
    powercfg /batteryreport /output "%DESK%\battery_report_%TS%.html" >nul 2>&1
    echo.
    echo ===== 15. PIN =====
    echo Da tao bao cao pin: %DESK%\battery_report_%TS%.html
)

:: ===== 16. Kiem tra toan ven file he thong (tuy chon) =====
echo.
echo ------------------------------------------------------------
choice /c YN /t 20 /d N /m "Chay them SFC va DISM CheckHealth (mat khoang 5-15 phut)"
if errorlevel 2 goto :finish

echo.
echo ===== SFC /verifyonly =====
sfc /verifyonly
echo.
echo ===== DISM /CheckHealth =====
DISM /Online /Cleanup-Image /CheckHealth > "%TEMP%\dism_check.txt"
type "%TEMP%\dism_check.txt"
echo. >> "%REPORT%"
echo ===== DISM /CheckHealth ===== >> "%REPORT%"
type "%TEMP%\dism_check.txt" >> "%REPORT%"
del "%TEMP%\dism_check.txt" >nul 2>&1

:finish
echo.
echo ============================================================
echo  HOAN TAT. Bao cao da luu tai:
echo  %REPORT%
echo ============================================================
start "" notepad "%REPORT%"
pause
exit /b

:: ===== Ham chay lenh PowerShell, hien thi + ghi vao bao cao =====
:ps
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $o = & { Write-Output '===== %~1 ====='; %~2 } | Out-String -Width 220; Write-Host $o; Add-Content -LiteralPath $env:REPORT -Value $o -Encoding UTF8"
exit /b
