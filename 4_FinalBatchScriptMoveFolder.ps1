Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    Multiselect = $false # Multiple files can be chosen
    Title = "Select user file"
	Filter = 'TXT (*.txt)|*.txt' # Specified file types
}

[void]$FileBrowser.ShowDialog()
$UserFile = $FileBrowser.FileName;

If($FileBrowser.FileNames -like "*\*") {
	# Do something 
	$FileBrowser.FileName #Lists selected files (optional)
}
else {
    Write-Host "Cancelled by user"
}

$merge = @{
    "Inkorg" = "Inkorgen";
    "Calendar" = "Kalender";
    "Sent Items" = "Skickat";
    "Drafts" = "Utkast";
    "Contacts" = "Kontakter";
    "Notes" = "Anteckningar";
}

$credentials = Get-Credential
$ewsurl = "https://outlook.office365.com/EWS/Exchange.asmx"

foreach ($User in Get-Content $UserFile) {
    .\Merge-MailboxFolder.ps1 -SourceMailbox $User -MergeFolderList $merge -ProcessSubfolders -Delete -Impersonate -EwsUrl $ewsurl -Credentials $credentials
}

Write-Host "Complete!" -ForegroundColor Green