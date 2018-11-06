#Run this job after your tape job.  Go into the tape job properties, Options>Advanced>Advanced tab.  Check 'run the following script after the job'
#change path of powershell script to match where you put it
#C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe  -Noninteractive -File "C:\PSScripts\TapeReport.ps1"




#Add Veeam Snap In
Add-PSSnapin -Name VeeamPSSnapIn

#Mail Server Variables, change these
$smtpServer = "exchange.contoso.com"
$smtpFrom = "veeam@contoso.com"
$smtpTo = "bgates@contoso.com"

#Change job name to match yours
$jobName = "GFS Tape Backup"
$job=get-vbrtapejob -name $jobName
$Session = [veeam.backup.core.cbackupsession]::GetByJob($job.id) | Sort CreationTime -Descending | select -First 1
[xml]$xml=$session.AuxData
$session1=$xml.TapeAuxData.TapeMediums.TapeMedium.name
$MediaSet = Get-VBRTapeMedium -name $session1 | Select-Object -Property MediaSet -Last 1| ft -hidetableheaders| Out-String;
$MediaSet2 = $MediaSet.subString(0,$MediaSet.length-14) | Out-String;

$MediaSetExpiration = Get-VBRTapeMedium -name $session1 | Select-Object -Property ExpirationDate -Last 1| ft -hidetableheaders| Out-String;
$MediaSet2Expiration = $MediaSetExpiration.subString(0,$MediaSetExpiration.length-17)| Out-String;


[xml]$xml=$session.AuxData
$session1=$xml.TapeAuxData.TapeMediums.TapeMedium.name

#Veeam Tape Variables
$x=(get-date).adddays(-4)

$content=Get-VBRTapeMedium -name $Session1 |select Name,LastWriteTime,ExpirationDate,MediaSet |Sort-Object MediaSet,ExpirationDate,Name -Descending

$tapes = $content |  
     Sort-Object LastWriteTime -Descending | 
     Select-Object Name, @{N='MediaSet';E={$_.MediaSet -replace "\s\d{1,2}\:\d{1,2}\s[AP]M"}}, @{N='ExpirationDate';E={$_.ExpirationDate.ToString('MM-dd-yyyy')}} 
     


$HTMLStyle = @"
<style>
BODY{font-family: Calibri; font-size: 9pt;}
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 6px;border-style: solid;border-color: black;background-color: #63C1ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
</style>
"@

$Header = @"
<table>
<colgroup><col/><col/><col/></colgroup>
<tr><th>Name</th><th>MediaSet</th><th>ExpirationDate</th></tr>
"@

$footer = "</table>"


#Setup new arrays for tape-set lists
$WeeklySetArray = New-Object System.Collections.ArrayList
$MonthlySetArray = New-Object System.Collections.ArrayList
$YearlySetArray = New-Object System.Collections.ArrayList
#clear any existing values from array
$WeeklySetArray.Clear()
$MonthlySetArray.Clear()
$YearlySetArray.Clear()


$WeeklySetTapes = $tapes | where {$_.mediaset -like "*weekly*"} | Sort-Object -Property expirationdate -Descending | Select-Object -First 1
$WeeklySetTapesCount = 0
$MonthlySetTapes = $tapes | where {$_.mediaset -like "*monthly*"} | Sort-Object -Property expirationdate -Descending | Select-Object -First 1
$MonthlySetTapesCount = 0
$YearlySetTapes = $tapes | where {$_.mediaset -like "*yearly*"} | Sort-Object -Property expirationdate -Descending | Select-Object -First 1
$YearlySetTapesCount = 0

$WeeklySetArray.Add("<b>"+$jobname+"</b><br>"+$WeeklySetTapes.MediaSet+"<br><b>Expires:</b>"+$WeeklySetTapes.ExpirationDate)
$MonthlySetArray.Add("<b>"+$jobname+"</b><br>"+$MonthlySetTapes.MediaSet+"<br><b>Expires:</b>"+$MonthlySetTapes.ExpirationDate)
$YearlySetArray.Add("<b>"+$jobname+"</b><br>"+$YearlySetTapes.MediaSet+"<br><b>Expires:</b>"+$YearlySetTapes.ExpirationDate)

$WeeklySetArray.Add($Header)
$MonthlySetArray.Add($Header)
$YearlySetArray.Add($Header)



foreach ($tape in $tapes){
if ($tape.mediaset -like "*weekly*"){
$WeeklySetArray.Add("<tr><td>"+$tape.name+"</td><td>"+$tape.mediaset+"</td><td>"+$tape.expirationdate+"</td></tr>`r`n")
$WeeklySetTapesCount++
}
elseif ($tape.mediaset -like "*monthly*"){
$MonthlySetArray.Add("<tr><td>"+$tape.name+"</td><td>"+$tape.mediaset+"</td><td>"+$tape.expirationdate+"</td></tr>`r`n")
$MonthlySetTapesCount++
}
elseif ($tape.mediaset -like "*yearly*"){
$YearlySetArray.Add("<tr><td>"+$tape.name+"</td><td>"+$tape.mediaset+"</td><td>"+$tape.expirationdate+"</td></tr>`r`n")
$YearlySetTapesCount++
}
}

$WeeklySetArray.Add($footer)
$MonthlySetArray.Add($footer)
$YearlySetArray.Add($footer)

if ($YearlySetTapesCount -eq "0"){$YearlySetArray.Clear()}


#Message Variables
$messageSubject = $MediaSet2 -Replace "[^ -~]", ""
$message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
$message.Subject = $messageSubject
$message.IsBodyHTML = $true

#Message Body

$bodyline1 = "<DIV style=font-size:11pt;font-family:Calibri> <p>The following tapes are ready to be sent offsite. Please replace with free ones:<br><br></p></DIV>"
$message.Body= $HTMLStyle+$bodyline1+$WeeklySetArray + "<br>"+$monthlySetArray

#Send the Message
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)
