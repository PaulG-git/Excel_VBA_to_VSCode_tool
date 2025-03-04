param (
  [string]$ExcelFileName = "your_excel_file.xlsm",
  [string]$ExcelFilePath = "C:\path_to_the_folder_your_file_is_in\$ExcelFileName",
  [string]$ScriptsFolder = "C:\path_to_folder_with_your_scripts"
)
# Get the workspace folder by moving up one directory from this script's location
$workspaceFolder = Split-Path -Parent $PSScriptRoot

# Construct the full paths for the scripts
$exportScript = Join-Path $workspaceFolder "\PS_Scripts\Export-VBA.ps1"
$importScript = Join-Path $workspaceFolder "\PS_Scripts\Import-VBA.ps1"

# Debugging Outputs
Write-Host "`n=== DEBUG INFO ==="
Write-Host "Workspace Folder: $workspaceFolder"
Write-Host "Export Script Path: $exportScript"
Write-Host "Import Script Path: $importScript"
Write-Host "Excel File Path: $ExcelFilePath"
Write-Host "Scripts Folder: $ScriptsFolder"
Write-Host "==================="

Write-Host "`nChoose an option:"
Write-Host "  1) Export VBA"
Write-Host "  2) Import VBA"
Write-Host "  3) Exit"
$choice = Read-Host "Enter 1, 2, or 3"

if ($choice -eq '1') {
  Write-Host "Running Export..."
  & $exportScript -ExcelFileName "$ExcelFileName" -ExcelFilePath "$ExcelFilePath" -ExportFolder "$ScriptsFolder"
}
elseif ($choice -eq '2') {
  Write-Host "Running Import..."
  & $importScript -ExcelFileName "$ExcelFileName" -ExcelFilePath "$ExcelFilePath" -ImportFolder "$ScriptsFolder"
}
elseif ($choice -eq '3') {
  Write-Host "Exiting... No action will be taken."
  exit 0
}
else {
  Write-Host "Invalid option. Please enter 1, 2, or 3."
}