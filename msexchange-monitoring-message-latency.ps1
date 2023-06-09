﻿#Monitoring Message Latency on Exchange servers
#

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://EXCHANGE-SERVER/PowerShell/ -Authentication Kerberos
Import-PSSession $Session

#Local domain
$localDomain = "@bgelov.ru"

#Exchange servers
$server = "Exch1","Exch2","Excha3","Exch4"

#Notify from
$smtpFrom = "MessageLatency@bgelov.ru"

#Notify to
$smtpTo = "MessageLatency@bgelov.ru"

#Message subject
$messageSubject = "Big MessageLatency on "

#SMTP server
$smtpserver = "smtp.bgelov.ru"

#Check period in min
$minutes = 30


#Lattency-----
#Seconds
$messageLatency = '00:00:12'
#MB
$messageLatencyMB = 5
#RecipientCounts
$messageRecipientsCount = 45


foreach ($s in $server) {
    $result = $null
    $body = '<html>
<head>
<style>
body {
    font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
    border-collapse: collapse;
    width: 100%;
}
table, td, th {
    border: 1px solid black;
    padding:6px;
}
table {
    border-collapse: collapse;
    width: 500px;
}

th {
    height: 30px;
}
</style>
</head>
<body>
<p>Наблюдаются задержки в доставке почты за последние ' + $minutes + ' минут.<br>
Возможно, есть какие-либо проблемы:</p>'

    $table = "<table><tr style='background-color:#fafafa'><th>EventID</th><th>Timestamp</th><th>MB</th><th>Sender</th><th>MessageLatency</th><th>Recipients</th><th>MessageSubject</th></tr>"

    $result = Get-MessageTrackingLog -resultsize Unlimited -server $s -start (get-date).AddMinutes(-$minutes) | SELECT eventid, TIMESTAMP, @{Label="MB"; Expression={$_.TotalBytes/1024/1024}}, sender, MessageLatency, @{Label="RecipientsCount"; Expression={$_.recipients.Count}}, MessageSubject | where {($_.eventid = 'SEND') -and ($_.sender -like "*$localDomain") -and ($_.MessageLatency -gt $messageLatency) -and ($_.MB -lt $messageLatencyMB) -and ($_.RecipientsCount -lt $messageRecipientsCount)}
 
    foreach ($r in $result) {

        $table += '<tr><td>' + $r.EventId + '</td><td>' + $r.Timestamp + '</td><td>' + $r.MB + '</td><td>' + $r.Sender + '</td><td>' + $r.MessageLatency + '</td><td>' + $r.RecipientsCount + '</td><td>' + $r.MessageSubject + '</td></tr>'

    }

    $table += "</table>"

    $body += $table + "<p>Письмо сгенерировано автоматически.</p>
    </body></html>"

    
    if ($result) {
        send-mailmessage -from "$smtpFrom" -to "$smtpTo" -subject "$messageSubject $s" -smtpServer "$smtpserver" -Body $body -Encoding UTF8 -BodyAsHtml
    }


}

