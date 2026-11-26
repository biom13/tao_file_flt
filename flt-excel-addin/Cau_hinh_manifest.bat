@echo off
chcp 65001 >nul
title Cau hinh manifest cho Excel add-in FLT
echo ============================================================
echo   Thay YOUR-USERNAME / YOUR-REPO trong manifest.xml
echo ============================================================
echo.
echo Vi du: neu link GitHub Pages cua ban la
echo   https://nguyenvana.github.io/flt-addin/
echo thi GitHub username = nguyenvana, ten repository = flt-addin
echo.
set "GH_USER="
set "GH_REPO="
set /p "GH_USER=Nhap GitHub username: "
set /p "GH_REPO=Nhap ten repository: "
if not defined GH_USER goto :thieu
if not defined GH_REPO goto :thieu

set "MANIFEST=%~dp0manifest.xml"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$p=$env:MANIFEST; $t=[IO.File]::ReadAllText($p); $t=$t.Replace('YOUR-USERNAME',$env:GH_USER).Replace('YOUR-REPO',$env:GH_REPO); [IO.File]::WriteAllText($p,$t,(New-Object Text.UTF8Encoding $false))"
if errorlevel 1 goto :loi

echo.
echo Da cap nhat manifest.xml. Add-in se chay tai:
echo   https://%GH_USER%.github.io/%GH_REPO%/taskpane.html
echo.
echo Buoc tiep: upload lai manifest.xml len GitHub (neu da upload ban cu)
echo va bat GitHub Pages. Xem huong dan trong README.md.
echo.
pause
exit /b 0

:thieu
echo.
echo Ban chua nhap du thong tin. Chay lai file nay.
pause
exit /b 1

:loi
echo.
echo Khong cap nhat duoc manifest.xml. Hay mo file bang Notepad va thay tay
echo YOUR-USERNAME va YOUR-REPO.
pause
exit /b 1
