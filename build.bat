@echo off
setlocal
rem Build DocTabs.exe with the .NET Framework 4.x compiler that ships with Windows.
rem Output goes to stage\DocTabs.exe: the output file name IS the assembly name, and the
rem assembly name must match the "Assembly" value registered for the Office add-in (DocTabs).
cd /d "%~dp0"
if not exist stage mkdir stage

set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
    echo [ERR] csc.exe not found. Install .NET Framework 4.x.
    if not defined CI pause
    exit /b 1
)

"%CSC%" /nologo /codepage:65001 /target:winexe /optimize+ /out:stage\DocTabs.exe /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll DocTabs.cs
if errorlevel 1 (
    echo [ERR] compile failed
    if not defined CI pause
    exit /b 1
)

powershell -NoProfile -Command "if([Reflection.AssemblyName]::GetAssemblyName('stage\DocTabs.exe').Name -ne 'DocTabs'){Write-Host '[ERR] assembly name must be DocTabs, otherwise Word silently skips the add-in'; exit 1}"
if errorlevel 1 (
    if not defined CI pause
    exit /b 1
)

echo OK: stage\DocTabs.exe
if not defined CI pause
