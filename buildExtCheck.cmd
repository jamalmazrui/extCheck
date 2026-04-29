@echo off
rem ============================================================================
rem buildExtCheck.cmd  --  Build extCheck.exe from source
rem
rem Compiles extCheck.cs with the C# compiler that ships with .NET Framework
rem 4.8.1, embeds extCheck.ico into the resulting executable, and copies the
rem build outputs into a dist\ subfolder ready for packaging by Inno Setup.
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
rem Output:
rem   dist\extCheck.exe   -- the built executable, with embedded icon
rem   dist\extCheck.csv   -- the rule registry, copied from the source dir
rem   dist\extCheck.ico   -- copied because Inno Setup needs it at compile time
rem                          for SetupIconFile=; not required at runtime
rem   dist\ReadMe.htm     -- you produce these from ReadMe.md (e.g. via Pandoc)
rem   dist\ReadMe.md      -- documentation
rem   dist\license.htm    -- license
rem   dist\announce.htm   -- release notes
rem   dist\announce.md    -- release notes (markdown)
rem   dist\extCheck.cs    -- source (for transparency / GPL-style spirit)
rem   dist\extCheck.iss   -- installer script
rem   dist\buildExtCheck.cmd -- build script
rem ============================================================================

setlocal enableextensions

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
echo Using compiler: %sCsc%

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

rem -- Prepare dist\ ----------------------------------------------------------
if not exist dist mkdir dist

rem -- Compile ----------------------------------------------------------------
echo Compiling extCheck.cs ...
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
    /out:dist\extCheck.exe ^
    extCheck.cs
if errorlevel 1 (
    echo [ERROR] Compilation failed.
    exit /b 1
)

rem -- Copy auxiliary files into dist\ ----------------------------------------
echo Copying auxiliary files into dist\ ...
copy /Y extCheck.csv dist\extCheck.csv >nul
if errorlevel 1 exit /b %errorlevel%
rem extCheck.ico is copied into dist\ only because Inno Setup needs the
rem file alongside extCheck.iss at compile time (for SetupIconFile=).
rem At runtime the icon is embedded in extCheck.exe (via /win32icon above),
rem so the .ico does NOT need to be installed alongside the exe.
copy /Y extCheck.ico dist\extCheck.ico >nul
if errorlevel 1 exit /b %errorlevel%

echo(
echo Build complete. The runtime distribution is just two files:
echo   dist\extCheck.exe   (the icon is embedded)
echo   dist\extCheck.csv   (the rule registry, used by -rules)
echo(
echo To produce the installer (extCheck_setup.exe):
echo   1. Copy ReadMe.htm, license.htm, ReadMe.md, announce.md,
echo      announce.htm, extCheck.cs, extCheck.iss, and
echo      buildExtCheck.cmd into dist\
echo      (extCheck.ico is needed here at compile time only, not at
echo      runtime, because Inno Setup uses it for the wizard's icon.)
echo   2. Open extCheck.iss in Inno Setup and click Compile.
echo      The result is dist\extCheck_setup.exe.
endlocal
