@echo off
setlocal enabledelayedexpansion

set DOTNET_EXE=dotnet
where dotnet >nul 2>&1
if errorlevel 1 (
  if exist "%ProgramFiles%\dotnet\dotnet.exe" (
    set DOTNET_EXE="%ProgramFiles%\dotnet\dotnet.exe"
  ) else (
    call :install_dotnet
  )
)

call :check_sdk
if not defined SDK_OK (
  call :install_dotnet
  call :check_sdk
)

if not defined SDK_OK (
  echo .NET 8 SDK not found. Please install it and re-run.
  start "" "https://dotnet.microsoft.com/download/dotnet/8.0"
  exit /b 1
)

set ROOT=%~dp0
set GUI_CSProj=%ROOT%tools\ANB-RenameBatch-GUI\ANB.RenameBatch.GUI.csproj
set OUT_DIR=%ROOT%tools\ANB-RenameBatch-GUI\dist

%DOTNET_EXE% publish "%GUI_CSProj%" -c Release -o "%OUT_DIR%"
if errorlevel 1 (
  echo Build failed.
  exit /b 1
)

echo.
echo Build complete.
echo Run: "%OUT_DIR%\ANB.RenameBatch.GUI.exe"
pause
exit /b 0

:check_sdk
set SDK_OK=
for /f "tokens=1" %%S in ('%DOTNET_EXE% --list-sdks 2^>nul') do (
  echo %%S | findstr /b 8. >nul && set SDK_OK=1
)
exit /b 0

:install_dotnet
echo .NET 8 SDK not found.
where winget >nul 2>&1
if errorlevel 1 (
  echo Winget not found. Opening .NET 8 SDK download page...
  start "" "https://dotnet.microsoft.com/download/dotnet/8.0"
  exit /b 1
) else (
  echo Installing .NET 8 SDK with winget...
  winget install --id Microsoft.DotNet.SDK.8 --source winget
)
if exist "%ProgramFiles%\dotnet\dotnet.exe" set DOTNET_EXE="%ProgramFiles%\dotnet\dotnet.exe"
exit /b 0
