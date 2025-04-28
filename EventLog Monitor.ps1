$ServerName = "Server Name"
$MailTo = 'Mail@email.com'
$MailFrom = 'Automation@email.com'
$SMTPServer = "Smtpserver1.company.com"
 
while ($true) {
    #$startTime = Get-Date -Hour 0 -Minute 0 -Second 0 -Millisecond 0
    $startTime = (Get-Date).Addminutes(-5)
    $events = Get-EventLog -LogName 'Citrix Delivery Services' -After $startTime -ComputerName $ServerName
    $Text = "All the Citrix XML Services configured for farm Xenapp5 failed to respond to this XML Service transaction."
    $Text2 = "The specified Citrix XML Service could not be contacted and has been temporarily removed from the list of active services."
    $server = $false
    foreach ($event in $events) {
        if ($event.Message -match $Text -or $event.Message -match $Text2) {
            $server = $true 
            $Time = $event.TimeGenerated
        }
    }
    if ($server) {
        $body = "`n Hi Team, `n`n Errors logged in Citrix server at $Time. Please login to citrix portal and launch the application. If you're able to launch it successfully, please ignore the email."
        Send-MailMessage -to $MailTo -SmtpServer $SMTPServer -Body $body -From $MailFrom -Subject "Citrix Site Error Log Monitoring"
    }
    <#
    else {
        $body = "No relevant Citrix Delivery Services events found since $startTime."
        Send-MailMessage -to $MailTo -SmtpServer $SMTPServer -Body $body -From $MailFrom -Subject "XenApp5 Citrix Site Error Log Monitoring"
    }#>
    Start-Sleep 300
}
