# Import-VBA.ps1
param (
  [string]$ExcelFileName = "your_excel_file.xlsm",
  [string]$ExcelFilePath = "C:\path_to_the_folder_your_file_is_in\$ExcelFileName",
  [string]$ImportFolder = "C:\path_of_folder_to_import_scripts"
)

# Debugging Output
Write-Host "`n=== DEBUG INFO (Import) ==="
Write-Host "Excel File Path: $ExcelFilePath"
Write-Host "Import Folder: $ImportFolder"
Write-Host "=============================="

# Check for null or empty ImportFolder
if ([string]::IsNullOrEmpty($ImportFolder)) {
  Write-Host "ERROR: ImportFolder is null or empty."
  exit 1
}

# Ensure the ImportFolder exists
if (-not (Test-Path $ImportFolder)) {
  Write-Host "ERROR: Import Folder does not exist: $ImportFolder"
  exit 1
}

# Specify the file extensions to import
$fileExtensions = @("*.bas", "*.cls", "*.frm", "*.txt")

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

$vbe = $workbook.VBProject.VBComponents

try {
  # Loop through each file type
  foreach ($ext in $fileExtensions) {
    Write-Host "`nLooking for files with extension: $ext"
        
    # Get files with the current extension
    $files = Get-ChildItem -Path $ImportFolder -Filter $ext -Recurse

    # Import each file
    foreach ($file in $files) {
      Write-Host "Processing: $($file.FullName)"
      $componentName = [IO.Path]::GetFileNameWithoutExtension($file.Name)

      # Check for special components (ThisWorkbook, Sheet1, etc.)
      $targetComponent = $null
      switch -Regex ($componentName) {
        "^ThisWorkbook$" {
          $targetComponent = $vbe.Item("ThisWorkbook")
        }
        "^Sheet\d+$" {
          $targetComponent = $vbe.Item($componentName)
        }
      }

      # Handle special components
      if ($targetComponent) {
        Write-Host "Handling special component: $componentName"

        try {
          # Clear existing code
          $targetComponent.CodeModule.DeleteLines(1, $targetComponent.CodeModule.CountOfLines)
                    
          # Add new code from file
          $targetComponent.CodeModule.AddFromFile($file.FullName)
          Write-Host "Updated code for: $componentName"
        }
        catch {
          Write-Host "ERROR: Failed to update $componentName - $($_.Exception.Message)"
        }
        continue
      }

      # Handle regular components (e.g., Modules, Classes, UserForms)
      Write-Host "Handling standard component: $componentName"
      try {
        # Check if the component already exists, if so, remove it first
        $existingComponent = $vbe | Where-Object { $_.Name -eq $componentName }
        if ($existingComponent) {
          Write-Host "Removing existing component: $componentName"
          $vbe.Remove($existingComponent)
        }

        # Add the component
        Write-Host "Adding new component: $componentName"
        $vbe.Import($file.FullName)

      }
      catch {
        Write-Host "ERROR: Failed to import $($file.Name) - $($_.Exception.Message)"
      }
    }
  }

  # Save the workbook after all imports
  Write-Host "`nSaving the workbook..."
  $workbook.Save()

  Write-Host "Workbook saved successfully."

}
catch {
  Write-Host "ERROR: An unexpected error occurred - $($_.Exception.Message)"
}
finally {
  # Release Excel restrictions
  $excel.DisplayAlerts = $true
  $excel.AskToUpdateLinks = $true
  $excel.EnableEvents = $true
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

Write-Host "Import complete."