#Run this as a scheduled task daily.  Task Scheduler on veeam server, new task
#Run task as system, run whether user is logged on or not, run as highest priveleges.
#Triggers, daily at whatever time(I do 8AM).  
#Actions, start a program
#Program/script: C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
#Arguments(Change as needed): -file "C:\PSScripts\ExpiredTapeReport.ps1" -ExecutionPolicy Bypass

#Add Veeam Snap In
Add-PSSnapin -Name VeeamPSSnapIn

#Email server setup, change values below
$smtpServer = "exchange.contoso.com"
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
#Email structure
$msg.From = "veeam@contoso.com"
$msg.To.Add("bgates@contoso.com")
$msg.IsBodyHTML = $true

$Header = @"
<style>
BODY{font-family: Calibri; font-size: 9pt;}
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 6px;border-style: solid;border-color: black;background-color: #63C1ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
</style>
"@

$currentdate=Get-Date -Format 'MM-dd-yyyy'
write-host $currentdate

$content=Get-VBRTapeMedium | select Name,LastWriteTime,ExpirationDate,MediaSet, isexpired | where {$_.isexpired -eq $false}

$sortedContent = $content |  
     Sort-Object LastWriteTime -Descending | 
     Select-Object Name, @{N='MediaSet';E={$_.MediaSet -replace "\s\d{1,2}\:\d{1,2}\s[AP]M"}}, @{N='ExpirationDate';E={$_.ExpirationDate.ToString('MM-dd-yyyy')}} | where {$_.Expirationdate -like $currentdate} 
     
     
     

$sortedcontent.Count

$sortedContentHTML=$sortedContent| ConvertTo-Html -Head $Header

if ($sortedContent.Count -gt 0) {


$msg.subject = "VEEAM - Expired Media - "+$sortedContent.count+" tapes"
$msg.body = "The following tapes have expired and can be reused<br><br>"+$sortedContentHTML
#Sending email
$smtp.Send($msg)

}

else {
#no tapes expired today, do nothing
}
