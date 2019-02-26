#Prompt user for O365 Credential
$UserCredential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Import-PSSession $Session -DisableNameChecking

#Prompt user for input about scope on who to run script on
$ScopeInput = Read-Host "Change language on All(A) users or specific(S) users?"

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Textfile (*.txt)|*.txt'
}
$null = $FileBrowser.ShowDialog()

$Users = Get-Content $FileBrowser

#Prompt user for input about language choise
$LanguageInput = Read-Host "English(en) or Swedish(se) as regional setting for all users mailboxes?"

switch ($LanguageInput)
{
    en {$LanguageCode = 1033}
    english {$LanguageCode = 1033}
    se {$LanguageCode = 1053}
    swedish {$LanguageCode = 1053}
}

if($ScopeInput -eq "A")
{
    Get-Mailbox -Filter {Name -notlike '*discover*'} | Set-MailboxRegionalConfiguration -Language $LanguageCode -TimeZone "W. Europe Standard Time" -LocalizeDefaultFolderName
}elseif ($ScopeInput -eq "S") {
    foreach($User in $Users){
        Get-Mailbox -Filter {Name -like '*$User*'} | Set-MailboxRegionalConfiguration -Language $LanguageCode -TimeZone "W. Europe Standard Time" -LocalizeDefaultFolderName
    }
    
}

Write-Host "All done!"