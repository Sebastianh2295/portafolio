@echo off
title Portafolio PMO
rem ===================================================================
rem  Portafolio PMO - abre el Excel del portafolio (que esta en linea)
rem  directamente en Excel de escritorio. El panel del Portafolio se
rem  abre solo si activaste esa opcion en el panel.
rem  Para cambiar el enlace: borra el archivo Portafolio_ruta.txt
rem  que queda junto a este .bat y vuelve a ejecutarlo.
rem ===================================================================
set "CFG=%~dp0Portafolio_ruta.txt"
if not exist "%CFG%" call :configurar
if not exist "%CFG%" exit /b

powershell -NoProfile -ExecutionPolicy Bypass -Command "$r=(Get-Content -LiteralPath $env:CFG -Raw -Encoding UTF8).Trim(); if($r -match '^https?://'){ Start-Process ('ms-excel:ofe|u|' + $r) } else { Start-Process -FilePath $r }"
exit /b

:configurar
cls
echo.
echo   ==========  Portafolio PMO - configuracion (solo la primera vez)  ==========
echo.
echo   1. Abre el Excel del Portafolio en el navegador (Excel en linea).
echo   2. Ve a  Archivo  ^>  Informacion  ^>  "Copiar ruta de acceso".
echo      (tambien sirve el enlace de la barra de direcciones del navegador)
echo   3. Pegalo aqui con clic derecho y presiona Enter.
echo.
set "RUTA="
set /p "RUTA=  Enlace del Excel: "
if not defined RUTA exit /b
powershell -NoProfile -ExecutionPolicy Bypass -Command "$u=$env:RUTA.Trim().Trim([char]34); if($u -match '^https?://'){ $u=$u -replace '/:[a-z]:/r/','/'; $u=$u.Split('?')[0]; $u=([Uri]$u).AbsoluteUri }; [IO.File]::WriteAllText($env:CFG, $u, (New-Object Text.UTF8Encoding $false)); Write-Host ''; Write-Host ('  Guardado: ' + $u); Start-Sleep -Seconds 2"
exit /b
