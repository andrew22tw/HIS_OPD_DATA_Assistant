@echo off
echo ==========================================
echo   Lab Data Formatter - Build
echo ==========================================
echo.

set CSC=

echo Searching for C# compiler...

:: Method 1: .NET Framework 4.x (64-bit)
for /f "delims=" %%i in ('dir /s /b "%WINDIR%\Microsoft.NET\Framework64\v4*\csc.exe" 2^>nul') do set CSC=%%i
:: Method 2: .NET Framework 4.x (32-bit)
if "%CSC%"=="" for /f "delims=" %%i in ('dir /s /b "%WINDIR%\Microsoft.NET\Framework\v4*\csc.exe" 2^>nul') do set CSC=%%i
:: Method 3: .NET Framework 3.5
if "%CSC%"=="" for /f "delims=" %%i in ('dir /s /b "%WINDIR%\Microsoft.NET\Framework64\v3*\csc.exe" 2^>nul') do set CSC=%%i
if "%CSC%"=="" for /f "delims=" %%i in ('dir /s /b "%WINDIR%\Microsoft.NET\Framework\v3*\csc.exe" 2^>nul') do set CSC=%%i
:: Method 4: .NET Framework 2.0
if "%CSC%"=="" for /f "delims=" %%i in ('dir /s /b "%WINDIR%\Microsoft.NET\Framework64\v2*\csc.exe" 2^>nul') do set CSC=%%i
if "%CSC%"=="" for /f "delims=" %%i in ('dir /s /b "%WINDIR%\Microsoft.NET\Framework\v2*\csc.exe" 2^>nul') do set CSC=%%i
:: Method 5: Any csc.exe under Microsoft.NET
if "%CSC%"=="" for /f "delims=" %%i in ('dir /s /b "%WINDIR%\Microsoft.NET\csc.exe" 2^>nul') do set CSC=%%i
:: Method 6: Visual Studio Roslyn compiler
if "%CSC%"=="" for /f "delims=" %%i in ('dir /s /b "%ProgramFiles%\Microsoft Visual Studio\*\*\MSBuild\Current\Bin\Roslyn\csc.exe" 2^>nul') do set CSC=%%i
if "%CSC%"=="" for /f "delims=" %%i in ('dir /s /b "%ProgramFiles(x86)%\Microsoft Visual Studio\*\*\MSBuild\Current\Bin\Roslyn\csc.exe" 2^>nul') do set CSC=%%i
:: Method 7: csc on PATH
if "%CSC%"=="" for /f "delims=" %%i in ('where csc.exe 2^>nul') do set CSC=%%i

if not "%CSC%"=="" (
    echo Found: %CSC%
    echo Compiling...
    "%CSC%" /target:winexe /out:LabFormatter.exe /optimize+ /nologo LabFormatter.cs
    if %ERRORLEVEL%==0 (
        echo.
        echo ==========================================
        for %%A in (LabFormatter.exe) do echo   OK! LabFormatter.exe [%%~zA bytes]
        echo ==========================================
        echo   Double-click LabFormatter.exe to run.
        pause
        exit /b 0
    ) else (
        echo Compile failed, trying dotnet...
    )
)

:: Method 8: dotnet CLI
echo.
echo csc.exe not found, trying dotnet CLI...
where dotnet >nul 2>&1
if %ERRORLEVEL%==0 (
    echo Found dotnet, building with SDK...
    if not exist "LabFmt" mkdir LabFmt
    cd LabFmt
    if not exist "LabFmt.csproj" (
        echo ^<Project Sdk="Microsoft.NET.Sdk"^> > LabFmt.csproj
        echo   ^<PropertyGroup^> >> LabFmt.csproj
        echo     ^<OutputType^>WinExe^</OutputType^> >> LabFmt.csproj
        echo     ^<TargetFramework^>net48^</TargetFramework^> >> LabFmt.csproj
        echo     ^<UseWindowsForms^>true^</UseWindowsForms^> >> LabFmt.csproj
        echo     ^<AssemblyName^>LabFormatter^</AssemblyName^> >> LabFmt.csproj
        echo   ^</PropertyGroup^> >> LabFmt.csproj
        echo ^</Project^> >> LabFmt.csproj
    )
    copy /y ..\LabFormatter.cs . >nul
    dotnet build -c Release -o ..\
    cd ..
    if exist "LabFormatter.exe" (
        echo.
        echo ==========================================
        for %%A in (LabFormatter.exe) do echo   OK! LabFormatter.exe [%%~zA bytes]
        echo ==========================================
        pause
        exit /b 0
    )
    echo dotnet build failed, trying net6.0-windows...
    cd LabFmt
    echo ^<Project Sdk="Microsoft.NET.Sdk"^> > LabFmt.csproj
    echo   ^<PropertyGroup^> >> LabFmt.csproj
    echo     ^<OutputType^>WinExe^</OutputType^> >> LabFmt.csproj
    echo     ^<TargetFramework^>net6.0-windows^</TargetFramework^> >> LabFmt.csproj
    echo     ^<UseWindowsForms^>true^</UseWindowsForms^> >> LabFmt.csproj
    echo     ^<AssemblyName^>LabFormatter^</AssemblyName^> >> LabFmt.csproj
    echo   ^</PropertyGroup^> >> LabFmt.csproj
    echo ^</Project^> >> LabFmt.csproj
    dotnet publish -c Release -r win-x64 --self-contained -o ..\publish
    cd ..
    if exist "publish\LabFormatter.exe" (
        copy /y publish\LabFormatter.exe .
        echo.
        echo ==========================================
        for %%A in (LabFormatter.exe) do echo   OK! LabFormatter.exe [%%~zA bytes]
        echo ==========================================
        pause
        exit /b 0
    )
)

echo.
echo ==========================================
echo   Could not find any C# compiler!
echo.
echo   Please install ONE of these:
echo.
echo   Option A (recommended, small):
echo     .NET SDK: https://dotnet.microsoft.com/download
echo     Choose ".NET SDK x64" and install
echo     Then run this build.bat again
echo.
echo   Option B:
echo     Visual Studio Build Tools:
echo     https://visualstudio.microsoft.com/downloads/
echo     Scroll to "Tools for Visual Studio" section
echo ==========================================
pause
