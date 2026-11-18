@echo off
rem Build ClaudePalette.dll (net46: AutoCAD 2018-2024, net8: AutoCAD 2025-2026, net10: AutoCAD 2027+) va cai bundle tu nap cho MAY DEV.
rem Dong AutoCAD truoc neu bundle dang nap (DLL bi khoa). Ban phat hanh: installer\pack_setup.py -> Setup.exe.
setlocal
set DOTNET=C:\Autodesk\tools\dotnet\dotnet.exe
if not exist "%DOTNET%" set DOTNET=dotnet
set DOTNET_CLI_TELEMETRY_OPTOUT=1
set DOTNET_GENERATE_ASPNET_CERTIFICATE=false
set OUT=%~dp0installer\bin\Release
set B=%APPDATA%\Autodesk\ApplicationPlugins\ClaudeCAD.bundle
rem RestoreLockedMode = "--locked-mode" cua restore (dotnet build khong nhan co nay truc tiep)
"%DOTNET%" build "%~dp0installer\ClaudePalette.csproj" -c Release -nologo -v:q -p:RestoreLockedMode=true || exit /b 1
if not exist "%B%\Contents\net46" mkdir "%B%\Contents\net46"
if not exist "%B%\Contents\net8" mkdir "%B%\Contents\net8"
if not exist "%B%\Contents\net10" mkdir "%B%\Contents\net10"
copy /y "%OUT%\net46\ClaudePalette.dll" "%B%\Contents\net46\" >nul || exit /b 1
copy /y "%OUT%\net8.0-windows\ClaudePalette.dll" "%B%\Contents\net8\" >nul || exit /b 1
copy /y "%OUT%\net10.0-windows\ClaudePalette.dll" "%B%\Contents\net10\" >nul || exit /b 1
for %%D in ("%B%\Contents" "%B%\Contents\net46" "%B%\Contents\net8" "%B%\Contents\net10") do copy /y "%~dp0app_icon.ico" "%%~D\icon.ico" >nul
copy /y "%~dp0PackageContents.xml" "%B%\" >nul || exit /b 1
if exist "%B%\Contents\ClaudePalette.dll" del "%B%\Contents\ClaudePalette.dll"
rem Palette tim server.py/mcp.json theo HKCU\Software\ClaudeCad\Home -> tro ve ban dev nay
for %%I in ("%~dp0.") do reg add HKCU\Software\ClaudeCad /v Home /t REG_SZ /d "%%~fI" /f >nul || exit /b 1
echo Installed: %B%