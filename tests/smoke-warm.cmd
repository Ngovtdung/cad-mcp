@echo off
rem Build tests\SmokeWarm.exe (thu tay claude that, khong can AutoCAD). Xem chu thich dau SmokeWarm.cs.
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
set OUT=%TEMP%\cad-mcp-smoke-bin
if not exist "%OUT%" mkdir "%OUT%"
"%FW%\csc.exe" /nologo /codepage:65001 /platform:x64 /target:exe /out:"%OUT%\SmokeWarm.exe" ^
  /r:System.IO.Compression.dll /r:System.Xml.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll ^
  /r:System.Web.Extensions.dll /r:System.Xaml.dll ^
  /r:"%FW%\WPF\PresentationCore.dll" /r:"%FW%\WPF\PresentationFramework.dll" /r:"%FW%\WPF\WindowsBase.dll" ^
  "%~dp0..\Attachments.cs" "%~dp0SmokeWarm.cs" || exit /b 1
copy /y "%~dp0envprobe_stub.py" "%OUT%\" >nul
echo Built: %OUT%\SmokeWarm.exe
