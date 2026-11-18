@echo off
rem Build + run Attachments.cs tests outside AutoCAD. Never touches the bundle or the running AutoCAD.
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
set OUT=%TEMP%\cad-mcp-test
if not exist "%OUT%" mkdir "%OUT%"
"%FW%\csc.exe" /nologo /codepage:65001 /platform:x64 /target:exe /out:"%OUT%\AttachTests.exe" ^
  /r:System.IO.Compression.dll /r:System.Xml.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll ^
  /r:System.Web.Extensions.dll /r:System.Xaml.dll ^
  /r:"%FW%\WPF\PresentationCore.dll" /r:"%FW%\WPF\PresentationFramework.dll" /r:"%FW%\WPF\WindowsBase.dll" ^
  "%~dp0..\Attachments.cs" "%~dp0..\Hotkeys.cs" "%~dp0AttachTests.cs" || exit /b 1
python -I -B "%~dp0mkfixtures.py" "%OUT%\fixtures" || exit /b 1
"%OUT%\AttachTests.exe" "%OUT%\fixtures" "%~dp0.."
