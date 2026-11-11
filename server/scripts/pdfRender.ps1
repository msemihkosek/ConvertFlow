param (
  [Parameter(Mandatory=$true)][string]$InputPath,
  [Parameter(Mandatory=$true)][string]$OutputDir,
  [double]$Scale = 2.0
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType=WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime]
$null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType=WindowsRuntime]
$null = [Windows.UI.Colors, Windows.UI, ContentType=WindowsRuntime]

$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]

$asTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
})[0]

function Await-AsyncOp($WinRtTask, $ResultType) {
  $netTask = $asTaskGeneric.MakeGenericMethod($ResultType).Invoke($null, @($WinRtTask))
  $netTask.Wait(-1) | Out-Null
  return $netTask.Result
}

function Await-AsyncAction($WinRtTask) {
  $netTask = $asTaskAction.Invoke($null, @($WinRtTask))
  $netTask.Wait(-1) | Out-Null
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
  [System.IO.Directory]::CreateDirectory($OutputDir) | Out-Null
}

# Copy to local temp to avoid OneDrive lock / path issues
$tempPdf = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "cf_pdf_" + [System.Guid]::NewGuid().ToString() + ".pdf")
Copy-Item -LiteralPath $InputPath -Destination $tempPdf -Force

try {
  $storageFile = Await-AsyncOp ([Windows.Storage.StorageFile]::GetFileFromPathAsync($tempPdf)) ([Windows.Storage.StorageFile])
  $pdfDoc = Await-AsyncOp ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) ([Windows.Data.Pdf.PdfDocument])

  $pageCount = $pdfDoc.PageCount
  $padLen = [Math]::Max(3, $pageCount.ToString().Length)

  for ($i = 0; $i -lt $pageCount; $i++) {
    $page = $pdfDoc.GetPage([uint32]$i)
    try {
      $size = $page.Size
      $renderOpts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
      $renderOpts.DestinationWidth = [uint32][Math]::Max(1, [Math]::Round($size.Width * $Scale))
      $renderOpts.DestinationHeight = [uint32][Math]::Max(1, [Math]::Round($size.Height * $Scale))
      $renderOpts.BackgroundColor = [Windows.UI.Colors]::White

      $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
      try {
        Await-AsyncAction ($page.RenderToStreamAsync($stream, $renderOpts))
        $stream.Seek(0)

        $netStream = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream)
        $pageNumStr = ($i + 1).ToString().PadLeft($padLen, '0')
        $outFile = [System.IO.Path]::Combine($OutputDir, "sayfa_${pageNumStr}.png")

        $fs = [System.IO.File]::Create($outFile)
        try {
          $netStream.CopyTo($fs)
        } finally {
          $fs.Dispose()
          $netStream.Dispose()
        }
      } finally {
        $stream.Dispose()
      }
    } finally {
      $page.Dispose()
    }
  }

  Write-Output "PDF_RENDER_OK:$pageCount"
} finally {
  Remove-Item -LiteralPath $tempPdf -Force -ErrorAction SilentlyContinue
}
