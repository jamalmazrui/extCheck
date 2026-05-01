@echo off
rem ============================================================================
rem buildExtCheck.cmd  --  Build extCheck.exe from source
rem
rem Compiles extCheck.cs with the C# compiler that ships with .NET Framework
rem 4.8.1 and embeds extCheck.ico into the resulting executable. The exe is
rem written to the current directory (alongside the source file, the icon,
rem the rule registry CSV, and extCheck.iss).
rem
rem Requirements:
rem   .NET Framework 4.8.1 Developer Pack (provides csc.exe and the 4.8.1
rem   reference assemblies). Install from
rem   https://dotnet.microsoft.com/download/dotnet-framework/net481
rem
rem   The script auto-detects csc.exe in the standard install locations under
rem   %WINDIR%\Microsoft.NET\Framework64\v4.0.30319 (this is the 64-bit csc;
rem   it can target any /platform setting).
rem
rem Architecture:
rem   /platform:x64 -- extCheck is built 64-bit. Office COM automation
rem   requires this process and the installed Microsoft Office to share the
rem   same bitness; modern Office (Microsoft 365, Office 2019+, Office 2024)
rem   is 64-bit by default. If a user has 32-bit Office installed, change
rem   /platform:x64 to /platform:x86 below and rebuild.
rem
rem Build outputs (all in the current directory):
rem   extCheck.exe     -- the built executable, with embedded icon
rem
rem To produce the installer setup.exe:
rem   Open extCheck.iss in Inno Setup and click Compile.
rem   Inno Setup writes extCheck_setup.exe to the same directory.
rem ============================================================================

setlocal enableextensions
cd /d "%~dp0"

rem -- Locate csc.exe ---------------------------------------------------------
set sCsc=
for %%V in (v4.0.30319) do (
    if exist "%WINDIR%\Microsoft.NET\Framework64\%%V\csc.exe" (
        set sCsc=%WINDIR%\Microsoft.NET\Framework64\%%V\csc.exe
        goto :foundCsc
    )
)
:foundCsc
if "%sCsc%"=="" (
    echo [ERROR] Could not find csc.exe under %WINDIR%\Microsoft.NET\Framework64
    echo         Install the .NET Framework 4.8.1 Developer Pack from
    echo         https://dotnet.microsoft.com/download/dotnet-framework/net481
    exit /b 1
)
echo [INFO] Using compiler: %sCsc%

rem -- Verify the icon exists -------------------------------------------------
if not exist extCheck.ico (
    echo [ERROR] extCheck.ico not found in %CD%.
    echo         The icon file is required to embed into the exe.
    exit /b 1
)

rem -- Verify the rules CSV exists --------------------------------------------
if not exist extCheck.csv (
    echo [ERROR] extCheck.csv not found in %CD%.
    echo         The rule registry CSV is required for the -rules option.
    exit /b 1
)

rem -- Compile ----------------------------------------------------------------
echo [INFO] Compiling extCheck.cs ...
"%sCsc%" ^
    /target:exe ^
    /platform:x64 ^
    /optimize+ ^
    /nologo ^
    /win32icon:extCheck.ico ^
    /reference:System.dll ^
    /reference:System.Core.dll ^
    /reference:System.Windows.Forms.dll ^
    /reference:System.Drawing.dll ^
    /out:extCheck.exe ^
    extCheck.cs
if errorlevel 1 (
    echo [ERROR] Compilation failed.
    exit /b 1
)

echo(
echo [INFO] Build complete. extCheck.exe is in %CD% (icon embedded).
echo(
echo To produce the installer (extCheck_setup.exe):
echo   Open extCheck.iss in Inno Setup and click Compile.
echo   Inno Setup writes extCheck_setup.exe to %CD%.
endlocal
