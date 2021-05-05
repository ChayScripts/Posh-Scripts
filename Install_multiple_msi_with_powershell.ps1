#copy all your msi files into a folder.
#create a text file(use dir /b from cmd prompt, and get all the file names, like file1.msi, file2.msi) copy all file names to that text file. 
#Run below script.

$msi = gc .\filename.txt
foreach ($msifile in $msi) 
{
Start-Process -FilePath "$env:systemroot\system32\msiexec.exe" -ArgumentList "/i `"$msifile`" /qn /passive" -Wait
}
