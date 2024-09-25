##VB script to refresh active window/desktop
Option Explicit
Dim WSHShell, strDesktop
Set WSHShell = WScript.CreateObject("WScript.Shell") 
strDesktop = WSHShell.SpecialFolders("Desktop") 
WSHShell.AppActivate strDesktop
WSHShell.SendKeys "{F5}"
WScript.Quit
##VB script end

#copy above vb script to c:\users\userid\ folder.

for ($i=1; $i -le 100000; $i++){
Invoke-Command -ScriptBlock {cscript 'C:\Users\userid\DesktopRefresh.vbs'}
foreach ($one in (150..1))
{
cls
write-host Refreshing in $one seconds
sleep 1
}
}

######################################################################################
# Second script 

Clear-Host
Echo "Keep alive w/ scroll lock..."
$WinShell = New-Object -com "Wscript.Shell"
while ($true)
{
$WinShell.sendkeys("{SCROLLLOCK}")
Start-Sleep -Seconds 120
}

######################################################################################
#Third Script - It will keep your windows session alive and restart your machine at 7:00 AM.
#Note: Make sure you start the script after 12:00 AM on any day or change the time as needed. 
# If you run the script at 9 PM for example, (Get-Date) -ge $RebootTime condition results in true, as on that date, 9 PM, will be greater than 7 AM.
#while ($true) runs forever and consumes more cpu. 

for ($i = 1; $i -le 100000; $i++) {
    $RebootTime = [datetime]"7:00"
    if ((Get-Date) -ge $RebootTime) {
        Stop-Computer -Force
    }
    write-host "Keep alive w/ scroll lock..." -foreground yellow
    $WinShell = New-Object -com "Wscript.Shell"
    $WinShell.sendkeys("{SCROLLLOCK}")
    foreach ($one in (120..1)) {
        Clear-Host
        write-host Refreshing in $one seconds
        Start-Sleep 1
    }
}

######################################################################################
# Fourth Script
# If you run below posh script, it will simulate scroll lock key press every 120 seconds. But you will have a powershell console on the screen. If you close it, script will stop.
# If you do not want a console on your screen, save this posh script to a location in your machine.
# create a batch script as shown below and run that batch script at your login.

#Posh Script
Clear-Host
Echo "Keep alive w/ scroll lock..."
$WinShell = New-Object -com "Wscript.Shell"
while ($true)
{
$WinShell.sendkeys("{SCROLLLOCK}")
Start-Sleep -Seconds 120
}

#Batch script
# In below example, above posh script is saved to C:\scripts folder.
# after you set the below script at logon, when you login to your machine, below script starts and runs above powershell script which will simulate scroll lock key press every 120 seconds.

TITLE Keep Alive
%SystemRoot%\system32\WindowsPowerShell\v1.0\PowerShell.exe -WindowStyle Hidden -noprofile -file "C:\Scripts\KeepAlive.ps1"
