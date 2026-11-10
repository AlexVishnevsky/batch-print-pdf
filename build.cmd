@echo off
rem Build without SDK: .NET Framework 4 csc + WinRT union metadata (Windows built-in PDF renderer)
cd /d "%~dp0"
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
set WM=C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.26100.0\Windows.winmd
"%FW%\csc.exe" /nologo /codepage:65001 /target:winexe /platform:anycpu /optimize+ /win32icon:app.ico /out:BatchPrint.exe ^
 /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Management.dll ^
 /r:"%FW%\System.Runtime.dll" /r:"%FW%\System.Runtime.WindowsRuntime.dll" /r:"%FW%\System.Threading.Tasks.dll" /r:"%FW%\System.Runtime.InteropServices.WindowsRuntime.dll" ^
 /r:"%WM%" ^
 BatchPrint.cs
