# Export-VBA.ps1
param (
  [string]$ExcelFileName = "your_excel_file.xlsm",
  [string]$ExcelFilePath = "C:\path_to_the_folder_your_file_is_in\$ExcelFileName",
  [string]$ExportFolder = "C:\path_of_folder_to_export_scripts"
)

# Debugging Output
Write-Host "`n=== DEBUG INFO (Export) ==="
Write-Host "Excel File Path: $ExcelFilePath"
Write-Host "Export Folder: $ExportFolder"
Write-Host "=============================="

# Create Export Directory if it doesn't exist
if (!(Test-Path $ExportFolder)) {
  New-Item -ItemType Directory -Path $ExportFolder
}

# Check if workbook is open and set object
$excelprocess = Get-Process | Where-Object { $_.ProcessName -eq "EXCEL" }
if (!($excelprocess -and (Get-Process | Where-Object { $_.MainWindowTitle -eq "$ExcelFileName - EXCEL" }))) {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.EnableEvents = $false
  $workbook = $excel.Workbooks.Open($ExcelFilePath)
}
else {
  $excel = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.EnableEvents = $false
  foreach ($wb in $excel.Workbooks) {
    if ($wb.Name -eq "$ExcelFileName") {
      $workbook = $wb
    }
  }
}

# Get VBProject
$vbProj = $workbook.VBProject

# Export Modules Sorted by Component Type
foreach ($component in $vbProj.VBComponents) {
  $componentName = $component.Name
  $componentType = $component.Type
  $fileExtension = ""
  $subfolder = ""

  switch ($componentType) {
    1 { 
      $fileExtension = ".bas"
      $subfolder = "Modules"
    }
    2 { 
      $fileExtension = ".cls"
      $subfolder = "Classes"
    }
    3 { 
      $fileExtension = ".frm"
      $subfolder = "Forms"
    }
    100 { 
      $fileExtension = ".bas"
      $subfolder = "DocumentModules"
    }
  }

  # Create Subfolder if it doesn't exist
  $targetFolder = Join-Path $ExportFolder $subfolder
  if (!(Test-Path $targetFolder)) {
    New-Item -ItemType Directory -Path $targetFolder
  }

  $exportPath = Join-Path $targetFolder "$componentName$fileExtension"
  $component.Export($exportPath)
  Write-Host "Exported $componentName to $exportPath"
  if ($componentName -eq "ThisWorkbook") {
    $my_file = $exportPath
    (Get-Content $my_file | Select-Object -Skip 4) | Set-Content $my_file
  }
}

# Release Excel restrictions
$excel.DisplayAlerts = $true
$excel.AskToUpdateLinks = $true
$excel.EnableEvents = $true
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Remove-Variable excel
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "Export complete."