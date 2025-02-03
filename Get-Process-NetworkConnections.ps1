# Enter Windows process name below
$processName = "chrome"
while ($true) {
    Get-Process -Name $processName | ForEach-Object {
        $id = $_.Id
        Write-Host "`nConnections for $processName process $id"
        Get-NetTCPConnection | Where-Object { $_.OwningProcess -eq $id } | 
        Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State | 
        Format-Table -AutoSize
    }
    
    Start-Sleep -Seconds 2
    
}

