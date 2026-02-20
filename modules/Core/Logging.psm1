Set-StrictMode -Version Latest

function Write-UiLog {
  param(
    [Parameter(Mandatory=$true)][System.Windows.Forms.TextBox]$LogTextBox,
    [Parameter(Mandatory=$true)][string]$Message
  )

  if ($LogTextBox.InvokeRequired) {
    $appendAction = [System.Action[System.Windows.Forms.TextBox,string]]{
      param($tb, $msg)
      Write-UiLog -LogTextBox $tb -Message $msg
    }
    [void]$LogTextBox.BeginInvoke($appendAction, @($LogTextBox, $Message))
    return
  }

  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $LogTextBox.AppendText("[$ts] $Message`r`n")
  $LogTextBox.SelectionStart = $LogTextBox.TextLength
  $LogTextBox.ScrollToCaret()
}

Export-ModuleMember -Function Write-UiLog
