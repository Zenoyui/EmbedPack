@echo off
setlocal EnableExtensions EnableDelayedExpansion
pushd "%~dp0" >nul

set "PROJECT_NAME=EmbedPack"
set "BUILD_DIR=build"

if not exist "CMakeLists.txt" (
  echo [ERROR] CMakeLists.txt not found: %CD%
  popd >nul
  exit /b 1
)

if not exist "%BUILD_DIR%" mkdir "%BUILD_DIR%" >nul 2>&1

echo [INFO] Configuring (default generator, x64)
cmake -S . -B "%BUILD_DIR%" -A x64
if errorlevel 1 (
  echo [ERROR] CMake configure failed.
  popd >nul
  exit /b 1
)

echo [INFO] Building (Release)
cmake --build "%BUILD_DIR%" --config Release
if errorlevel 1 (
  echo [ERROR] Build failed.
  popd >nul
  exit /b 1
)

set "OUT_EXE=%BUILD_DIR%\Release\%PROJECT_NAME%.exe"
if not exist "%OUT_EXE%" (
  echo [ERROR] Output not found: "%OUT_EXE%"
  popd >nul
  exit /b 1
)

echo [OK] Release build completed.
echo [OK] Output: "%CD%\%OUT_EXE%"

popd >nul
exit /b 0
