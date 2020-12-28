#copy all the urls and servers that has certificate installed in your environment into text file rename it as urls.txt
#After you run the script, in same file, script will append data by adding date, time, url and server cert info.

$urls = Get-Content $psscriptroot\urls.txt 
$date = get-date -Format G 
$finaloutput = @() 
foreach ($url in $urls) { 
    [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }  
    $req = [Net.HttpWebRequest]::Create($url)
    try { 
        $req.GetResponse() | Out-Null 
 
        $req.ServicePoint.Certificate.GetIssuerName().split(",", [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -match "O=" } | Out-Null 
        if ($LASTEXITCODE) { 
            $certissuer = ($req.ServicePoint.Certificate.GetIssuerName().split(",", [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -match "O=" }).trim("O =")   
        }
        else { 
 
            $certissuer = ($req.ServicePoint.Certificate.GetIssuerName().split(",", [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object { $_ -match "CN=" } ).trim("CN =") 
        } 
 
        $output = [PSCustomObject]@{  
            URL               = $url  
            'Cert Start Date' = $req.ServicePoint.Certificate.GetEffectiveDateString()  
            'Cert End Date'   = $req.ServicePoint.Certificate.GetExpirationDateString() 
            'Cert Issuer'     = $certissuer 
            "Today's Date"    = $date 
     
        } 
    } 
    catch { 
        $formatteddata = [PSCustomObject]@{ 
            URL = "Cant get cert details for $url" 
        } 
    }
    $finaloutput += $output
} 
$finaloutput += $formatteddata
$finaloutput | export-csv -Path $psscriptroot\CertDetails.csv -Append -NoTypeInformation
