#Prompt user for O365 Credential
$UserCredential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

Import-PSSession $Session -DisableNameChecking

#Ask user for input about language choise
$LanguageInput = Read-Host "English(en) or Swedish(se) as regional setting for all users mailboxes?"

switch ($LanguageInput)
{
    en {$LanguageCode = 1033}
    english {$LanguageCode = 1033}
    se {$LanguageCode = 1053}
    swedish {$LanguageCode = 1053}
}

Get-Mailbox -Filter {Name -notlike '*discover*'} | Set-MailboxRegionalConfiguration -Language $LanguageCode -TimeZone "W. Europe Standard Time" -LocalizeDefaultFolderName