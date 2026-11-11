param (
  [Parameter(Mandatory=$true)][string]$Title,
  [Parameter(Mandatory=$true)][string]$Message,
  [switch]$IsError
)

$ErrorActionPreference = 'SilentlyContinue'

try {
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime] | Out-Null
  [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType=WindowsRuntime] | Out-Null

  $escapedTitle = [System.Security.SecurityElement]::Escape($Title)
  $escapedMsg = [System.Security.SecurityElement]::Escape($Message)

  $toastXmlStr = @"
<toast duration="short">
  <visual>
    <binding template="ToastGeneric">
      <text>$escapedTitle</text>
      <text>$escapedMsg</text>
    </binding>
  </visual>
</toast>
"@

  $xmlDoc = New-Object Windows.Data.Xml.Dom.XmlDocument
  $xmlDoc.LoadXml($toastXmlStr)

  $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
  $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId)
  $toast = [Windows.UI.Notifications.ToastNotification]::new($xmlDoc)
  $notifier.Show($toast)
} catch {
  if ($IsError) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show($Message, $Title, 'OK', 'Error') | Out-Null
  }
}

