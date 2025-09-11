@echo off
echo Fire Alarm Circuit Analysis - Deployment Script
echo ================================================

set REVIT_VERSION=2024
set ADDIN_PATH=C:\ProgramData\Autodesk\Revit\Addins\%REVIT_VERSION%
set PROJECT_PATH=%~dp0
set OUTPUT_PATH=%PROJECT_PATH%bin\Release

echo Building Release configuration...
msbuild FireAlarmCircuitAnalysis.csproj /p:Configuration=Release /p:Platform=AnyCPU

if %ERRORLEVEL% NEQ 0 (
    echo Build failed!
    pause
    exit /b 1
)

echo Creating deployment directories...
if not exist "%ADDIN_PATH%\FireAlarmCircuitAnalysis" mkdir "%ADDIN_PATH%\FireAlarmCircuitAnalysis"

echo Copying files...
xcopy "%OUTPUT_PATH%\FireAlarmCircuitAnalysis.dll" "%ADDIN_PATH%\FireAlarmCircuitAnalysis\" /Y
xcopy "%OUTPUT_PATH%\*.json" "%ADDIN_PATH%\FireAlarmCircuitAnalysis\" /Y
xcopy "%OUTPUT_PATH%\Newtonsoft.Json.dll" "%ADDIN_PATH%\FireAlarmCircuitAnalysis\" /Y
xcopy "%OUTPUT_PATH%\NPOI.dll" "%ADDIN_PATH%\FireAlarmCircuitAnalysis\" /Y
xcopy "%OUTPUT_PATH%\NPOI.OOXML.dll" "%ADDIN_PATH%\FireAlarmCircuitAnalysis\" /Y
xcopy "%OUTPUT_PATH%\NPOI.OpenXml4Net.dll" "%ADDIN_PATH%\FireAlarmCircuitAnalysis\" /Y
xcopy "%OUTPUT_PATH%\NPOI.OpenXmlFormats.dll" "%ADDIN_PATH%\FireAlarmCircuitAnalysis\" /Y
xcopy "%OUTPUT_PATH%\ICSharpCode.SharpZipLib.dll" "%ADDIN_PATH%\FireAlarmCircuitAnalysis\" /Y
xcopy "%OUTPUT_PATH%\NLog.dll" "%ADDIN_PATH%\FireAlarmCircuitAnalysis\" /Y
xcopy "%OUTPUT_PATH%\FireAlarmCircuitAnalysis.addin" "%ADDIN_PATH%\" /Y

echo Deployment complete!
echo Plugin installed to: %ADDIN_PATH%
pause