CD /d "%~dp0"
powershell set-executionpolicy bypass -force
powershell start-process powershell -verb runas %~dp0FolderScanningTool.ps1


