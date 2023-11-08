$OutputFileName = "$ENV:Workspace\NoLogonAccounts.csv"
$Day = (Get-Date).DayOfWeek

Import-Module activedirectory
$Days = (get-date).AddDays(-7)
Get-Aduser -Properties * -f {-not ( lastlogontimestamp -like "*") -and (enabled -eq $true) -and (aDAccountType -eq 'User')} | Where-object {$_.whencreated -lt $Days} | Select -Property Samaccountname,displayname,enabled,lastlogondate,lastlogontimestamp,whencreated,wcCurrentHireDate | export-csv -NoTypeInformation $OutputFileName


$Splat = @{ 

    To = "to@mail.com"
    From = "ADREporting@ADReporting.com"
    Body = "For your Review"
    Attachments = $OutputFileName
    Subject =  "No Logon Accounts Report $day" 
    SmtpServer = 'AM1SMTP' 
    BodyAsHtml = $True
}

Send-MailMessage @Splat 
