@echo off
chcp 65001 > nul
title Classic Sports - Desinstalador

echo.
echo  =====================================================
echo   CLASSIC SPORTS - Desinstalador
echo  =====================================================
echo.
echo  Esta acao vai remover o Classic Sports
echo  do seu computador.
echo.
echo  Seus relatorios serao salvos em:
echo  %USERPROFILE%\Documents\ClassicSports_Relatorios
echo.
set /p CONFIRM=Digite SIM para confirmar a desinstalacao: 
if /i not "%CONFIRM%"=="SIM" (
    echo Desinstalacao cancelada.
    pause
    exit /b 0
)

echo.
echo [1/5] Encerrando o programa...
taskkill /f /im pythonw.exe > nul 2>&1
taskkill /f /im python.exe > nul 2>&1
timeout /t 2 /nobreak > nul

echo [2/5] Salvando seus relatorios em Documentos...
set "PASTA_RELAT=%USERPROFILE%\Documents\ClassicSports_Relatorios"
if not exist "%PASTA_RELAT%" mkdir "%PASTA_RELAT%"
xcopy /E /Y /Q "C:\ClassicSportsApps\Relatorios\relatorios_*" "%PASTA_RELAT%\" > nul 2>&1
echo [OK] Relatorios salvos em: %PASTA_RELAT%

echo [3/5] Removendo atalhos...
:: Le caminho exato do atalho do registro do Windows
for /f "tokens=3*" %%A in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClassicSports" /v ShortcutPath 2^>nul') do set "SHORTCUT_PATH=%%A %%B"
if defined SHORTCUT_PATH (
    del /f /q "%SHORTCUT_PATH%" > nul 2>&1
    powershell -NoProfile -Command "Remove-Item -Path '%SHORTCUT_PATH%' -Force -ErrorAction SilentlyContinue"
)
:: Remove da area de trabalho do usuario atual (fallback)
del /f /q "%USERPROFILE%\Desktop\Classic Sports.lnk" > nul 2>&1
:: Remove da area de trabalho publica
del /f /q "%PUBLIC%\Desktop\Classic Sports.lnk" > nul 2>&1
:: Remove do menu iniciar
del /f /q "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Classic Sports.lnk" > nul 2>&1
del /f /q "%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs\Classic Sports.lnk" > nul 2>&1
:: Remove via PowerShell para garantir (trata caminhos com espacos)
powershell -NoProfile -Command "Remove-Item -Path '%USERPROFILE%\Desktop\Classic Sports.lnk' -Force -ErrorAction SilentlyContinue"
powershell -NoProfile -Command "Remove-Item -Path '%PUBLIC%\Desktop\Classic Sports.lnk' -Force -ErrorAction SilentlyContinue"
echo [OK]

echo [4/5] Removendo do registro do Windows...
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClassicSports" /f > nul 2>&1
reg delete "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\ClassicSports" /f > nul 2>&1
echo [OK]

echo [5/5] Removendo arquivos do programa...
:: Remove a pasta principal (mas este bat ja foi copiado para temp)
set "TEMP_BAT=%TEMP%\cs_uninstall_final.bat"
echo @echo off > "%TEMP_BAT%"
echo timeout /t 2 /nobreak ^> nul >> "%TEMP_BAT%"
echo rmdir /s /q "C:\ClassicSportsApps\Relatorios" >> "%TEMP_BAT%"
echo echo. >> "%TEMP_BAT%"
echo echo  Classic Sports removido com sucesso! >> "%TEMP_BAT%"
echo pause >> "%TEMP_BAT%"
start /min cmd /c "%TEMP_BAT%"
echo [OK]

echo.
echo  =====================================================
echo   Classic Sports sera removido em instantes!
echo   Seus relatorios foram salvos em:
echo   %PASTA_RELAT%
echo  =====================================================
echo.
pause
