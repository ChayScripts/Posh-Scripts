#you may have to change the url below when you use it. 
#Unregister-ScheduledTask is optional. 

$source = 'http://dl.delivery.mp.microsoft.com/filestreamingservice/files/63c21fa1-e10e-40ef-a9ac-b51f877c7cfb/MicrosoftEdgeEnterpriseX64.msi'
$destination = $env:temp + '\MSEdge.msi'
Invoke-WebRequest $source -OutFile $destination
Start-Process -FilePath $destination -ArgumentList "/qn /norestart" -wait
Unregister-ScheduledTask -TaskPath "\" -Confirm:$false
Remove-Item "$env:Temp\*" -recurse -force
