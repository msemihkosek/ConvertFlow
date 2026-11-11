param (
  [Parameter(Mandatory=$true)][string]$InputPath,
  [Parameter(Mandatory=$true)][string]$OutputPath,
  [Parameter(Mandatory=$true)][string]$TargetFormat
)

$ErrorActionPreference = 'Stop'

$ext = [System.IO.Path]::GetExtension($InputPath).ToLower().TrimStart('.')
$target = $TargetFormat.ToLower().Trim()

# Create a local temp working file outside any cloud-sync (OneDrive) folder
$tempWorkFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "cf_work_" + [System.Guid]::NewGuid().ToString() + "." + $ext)
Copy-Item -LiteralPath $InputPath -Destination $tempWorkFile -Force
try { Unblock-File -LiteralPath $tempWorkFile -ErrorAction SilentlyContinue } catch {}

try {
  # -------------------------------------------------------------
  # 1. PowerPoint Presentations (.ppt, .pptx, .pot, .pps, etc.)
  # -------------------------------------------------------------
  if ($ext -in @('ppt', 'pptx', 'pot', 'potx', 'pps', 'ppsx')) {
    $ppt = New-Object -ComObject PowerPoint.Application
    try {
      # Open(FileName, ReadOnly, Untitled, WithWindow)
      $pres = $ppt.Presentations.Open($tempWorkFile, -1, 0, 0)
      try {
        if ($target -eq 'pdf') {
          # 32 = ppSaveAsPDF
          $pres.SaveAs($OutputPath, 32)
        } elseif ($target -eq 'pptx') {
          # 24 = ppSaveAsOpenXMLPresentation
          $pres.SaveAs($OutputPath, 24)
        } elseif ($target -in @('png', 'jpg', 'jpeg', 'zip')) {
          $tempDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "cf_slides_" + [System.Guid]::NewGuid().ToString())
          [System.IO.Directory]::CreateDirectory($tempDir) | Out-Null
          $imgExt = if ($target -eq 'jpg' -or $target -eq 'jpeg') { 'jpg' } else { 'png' }
          $fmt = if ($target -eq 'jpg' -or $target -eq 'jpeg') { 'JPG' } else { 'PNG' }
          
          for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $slideFile = [System.IO.Path]::Combine($tempDir, ("Slide_{0:D3}.{1}" -f $i, $imgExt))
            $slide.Export($slideFile, $fmt, 1920, 1080)
          }

          if (Test-Path $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }
          Compress-Archive -Path "$tempDir\*" -DestinationPath $OutputPath -Force
          Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        } else {
          throw "PowerPoint hedef formatı desteklenmiyor: $target"
        }
      } finally {
        $pres.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
      }
    } finally {
      $ppt.Quit()
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
    }
  }
  # -------------------------------------------------------------
  # 2. Word Documents (.doc, .docx, .rtf, .dot, .dotx)
  # -------------------------------------------------------------
  elseif ($ext -in @('doc', 'docx', 'rtf', 'dot', 'dotx')) {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    try {
      $doc = $word.Documents.Open($tempWorkFile, $false, $true)
      try {
        if ($target -eq 'pdf') {
          # 17 = wdFormatPDF
          $doc.SaveAs2($OutputPath, 17)
        } elseif ($target -eq 'docx') {
          # 16 = wdFormatXMLDocument
          $doc.SaveAs2($OutputPath, 16)
        } elseif ($target -eq 'txt') {
          # 2 = wdFormatText
          $doc.SaveAs2($OutputPath, 2)
        } else {
          throw "Word hedef formatı desteklenmiyor: $target"
        }
      } finally {
        $doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
      }
    } finally {
      $word.Quit()
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
    }
  }
  # -------------------------------------------------------------
  # 3. Excel Spreadsheets (.xls, .xlsx, .xlsm, .csv)
  # -------------------------------------------------------------
  elseif ($ext -in @('xls', 'xlsx', 'xlsm', 'xlsb')) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
      $wb = $excel.Workbooks.Open($tempWorkFile, 0, $true)
      try {
        if ($target -eq 'pdf') {
          # 0 = xlTypePDF
          $wb.ExportAsFixedFormat(0, $OutputPath)
        } elseif ($target -eq 'xlsx') {
          # 51 = xlOpenXMLWorkbook
          $wb.SaveAs($OutputPath, 51)
        } elseif ($target -eq 'csv') {
          # 6 = xlCSV
          $wb.SaveAs($OutputPath, 6)
        } else {
          throw "Excel hedef formatı desteklenmiyor: $target"
        }
      } finally {
        $wb.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
      }
    } finally {
      $excel.Quit()
      [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
    }
  }
  else {
    throw "Desteklenmeyen dosya türü: $ext"
  }
} finally {
  Remove-Item -LiteralPath $tempWorkFile -Force -ErrorAction SilentlyContinue
}

Write-Output "CONVERSION_SUCCESS"
