add-type -AssemblyName Microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms

# $NovaLCTFilePath = "C:\Users\KLago\AppData\Roaming\Nova Star\NovaLCT\Bin\NovaLCT.exe"
# $StartNovaLCT = {Start-Process -FilePath $NovaLCTFilePath}
# $NovaLCTIsRunning = Get-Process -Name NovaLCT -ErrorAction 'silentlycontinue'
# $Test = New-Object -ComObject wscript.shell

# if ($NovaLCTIsRunning -eq $null) {
#    Write-Host "Attempting to start NovaLCT"
#    & $StartNovaLCT
#} else {
#    Write-Host "NovaLCT already running"
#}

#Microsoft.VisualBasic.Interaction.AppActivate("novalct");

[Microsoft.VisualBasic.Interaction]::AppActivate("NovaLCT");
#[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(50, 50);
Start-Sleep 1
[System.Windows.Forms.SendKeys]::SendWait("{TAB}")
# $Test::SendWait("s")