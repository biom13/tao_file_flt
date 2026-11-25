@echo off
setlocal EnableExtensions
title Tat hoan toan BitLocker - Windows 11

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
echo    TAT HOAN TOAN BITLOCKER - GIAI MA TAT CA O DIA
echo ============================================================
echo.
echo Hien trang BitLocker:
powershell -NoProfile -Command "Get-BitLockerVolume | Select-Object MountPoint,VolumeType,VolumeStatus,ProtectionStatus,LockStatus,EncryptionPercentage | Format-Table -AutoSize"
echo.
echo LUU Y:
echo  - Giai ma mat tu vai chuc phut den vai gio, tuy dung luong o dia.
echo  - Nen cam sac (laptop) va khong tat may khi dang giai ma.
echo    Neu lo tat, qua trinh se tu tiep tuc o lan khoi dong sau.
echo  - O nao dang bi KHOA can co Recovery Password 48 so de mo khoa.
echo  - Sau khi giai ma, du lieu tren o dia khong con duoc bao ve boi ma hoa.
echo.
choice /c YN /m "Ban co chac chan muon giai ma TAT CA o dia"
if errorlevel 2 goto :cancel

:: ===== Giai ma tung o dang bat BitLocker =====
set "FOUND=0"
for /f "usebackq delims=" %%M in (`powershell -NoProfile -Command "Get-BitLockerVolume | Where-Object { $_.VolumeStatus -ne 'FullyDecrypted' } | ForEach-Object { $_.MountPoint }"`) do (
    set "FOUND=1"
    call :decrypt "%%M"
)
if "%FOUND%"=="0" (
    echo.
    echo Khong co o dia nao dang bat BitLocker.
    goto :policy
)

:: ===== Theo doi tien trinh giai ma =====
:wait
cls
echo Dang giai ma cac o dia... (dong cua so nay giai ma van tiep tuc chay nen)
echo.
powershell -NoProfile -Command "Get-BitLockerVolume | Select-Object MountPoint,VolumeStatus,EncryptionPercentage,LockStatus | Format-Table -AutoSize"
set "PENDING=0"
for /f "usebackq delims=" %%N in (`powershell -NoProfile -Command "@(Get-BitLockerVolume | Where-Object { $_.VolumeStatus -eq 'DecryptionInProgress' -and $_.LockStatus -ne 'Locked' }).Count"`) do set "PENDING=%%N"
if not "%PENDING%"=="0" (
    timeout /t 10 /nobreak >nul
    goto :wait
)
echo.
echo Da xong buoc giai ma.

:: ===== Chan Windows tu bat lai ma hoa, tat dich vu BitLocker =====
:policy
echo.
echo ------------------------------------------------------------
echo Windows 11 co the TU DONG bat lai ma hoa (Device Encryption).
echo Buoc nay se:
echo   1. Dat chinh sach PreventDeviceEncryption = 1
echo   2. Tat dich vu BitLocker (BDESVC)
echo.
choice /c YN /m "Ap dung de BitLocker khong tu bat lai"
if errorlevel 2 goto :finish

reg add "HKLM\SYSTEM\CurrentControlSet\Control\BitLocker" /v PreventDeviceEncryption /t REG_DWORD /d 1 /f
sc config BDESVC start= disabled
sc stop BDESVC >nul 2>&1
echo Da chan BitLocker tu bat lai.

:finish
echo.
echo ============================================================
echo  Trang thai cuoi cung:
echo ============================================================
powershell -NoProfile -Command "Get-BitLockerVolume | Select-Object MountPoint,VolumeStatus,ProtectionStatus,LockStatus,EncryptionPercentage | Format-Table -AutoSize"
echo Neu con o nao o trang thai khac FullyDecrypted, do la o bi khoa
echo hoac giai ma chua hoan tat. Hay chay lai file nay.
echo.
pause
exit /b

:cancel
echo Da huy. Khong co thay doi nao duoc thuc hien.
pause
exit /b

:: ===== Ham giai ma mot o dia =====
:decrypt
set "MP=%~1"
set "LS="
for /f "usebackq delims=" %%L in (`powershell -NoProfile -Command "(Get-BitLockerVolume -MountPoint '%MP%').LockStatus"`) do set "LS=%%L"
if /i not "%LS%"=="Locked" goto :decrypt_go
echo.
echo O %MP% dang bi KHOA.
echo Nhap Recovery Password 48 so, hoac bam Enter de bo qua o nay:
set "RK="
set /p "RK=> "
if not defined RK goto :decrypt_skip
manage-bde -unlock %MP% -RecoveryPassword %RK%
if errorlevel 1 goto :decrypt_skip

:decrypt_go
echo.
echo Bat dau giai ma o %MP% ...
manage-bde -off %MP%
exit /b

:decrypt_skip
echo Bo qua o %MP% (van dang khoa hoac ma hoa).
exit /b
