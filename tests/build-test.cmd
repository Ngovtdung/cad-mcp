@echo off
rem Build ClaudePalette.dll into %TEMP%\cad-mcp-test for tests outside AutoCAD (no bundle install, no AutoCAD).
set ACAD=C:\Program Files\Autodesk\AutoCAD 2022
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
set OUT=%TEMP%\cad-mcp-test
if not exist "%OUT%" mkdir "%OUT%"
"%FW%\csc.exe" /nologo /codepage:65001 /platform:x64 /target:library /out:"%OUT%\ClaudePalette.dll" ^
  /resource:"%~dp0..\ClaudePalette.xaml",ClaudePalette.xaml ^
  /r:"%ACAD%\accoremgd.dll" /r:"%ACAD%\acdbmgd.dll" /r:"%ACAD%\acmgd.dll" /r:"%ACAD%\AdWindows.dll" ^
  /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Web.Extensions.dll /r:System.Xaml.dll ^
  /r:System.IO.Compression.dll /r:System.Xml.dll ^
  /r:"%FW%\WPF\PresentationCore.dll" /r:"%FW%\WPF\PresentationFramework.dll" /r:"%FW%\WPF\WindowsBase.dll" ^
  "%~dp0..\ClaudePalette.cs" "%~dp0..\Attachments.cs" "%~dp0..\Hotkeys.cs" || exit /b 1
echo Built: %OUT%\ClaudePalette.dll
