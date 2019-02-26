$DestinationURL = Read-Host "Please input destination URL to Azure"

Add-Type -AssemblyName System.Windows.Forms
$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
[void]$FolderBrowser.ShowDialog()
$PSTFolder = $FolderBrowser.SelectedPath

$confirmation = Read-Host "This will upload all files in $PSTFolder. Are you sure? [y/n]"
while($confirmation -ne "y")
{
    if ($confirmation -eq 'n') {exit}
    $confirmation = Read-Host "This will upload all files in $PSTFolder. Are you sure? [y/n]"
}

& '.\Program Files (x86)\Microsoft SDKs\Azure\AzCopy\AzCopy.exe' /Source:$PSTFolder /Dest:$DestinationURL /V:C:\ImportLog\UploadLog.log /Y