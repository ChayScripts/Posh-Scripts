#Requires -RunAsAdministrator
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.REflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$WinForm = new-object System.Windows.Forms.Form   
$WinForm.text = "Shortcut to files in Run dialog box"   
$WinForm.Size = new-object Drawing.Size(400,200)
$WinForm.StartPosition = "CenterScreen"
$WinForm.KeyPreview = $True
$WinForm.Add_KeyDown({ if ($_.KeyCode -eq "Escape") {$WinForm.Close()}})


$objTextBox1 = New-Object System.Windows.Forms.TextBox 
$objTextBox1.Location = New-Object System.Drawing.Size(10,20) 
$objTextBox1.Size = New-Object System.Drawing.Size(260,20) 
$winform.Controls.Add($objTextBox1) 


function Select-FileDialog
{
  param([string]$Filter="All Files (*.*)|*.*;*.*")
	$objForm = New-Object System.Windows.Forms.OpenFileDialog
	$objForm.InitialDirectory = $Directory
	$objForm.Filter = $Filter
	$objForm.Title = $Title
	$Show = $objForm.ShowDialog()
	If ($Show -eq "OK")
	{
		Return $objForm.FileName
	}
	
}

$SelectButton1 = New-Object System.Windows.Forms.Button
$SelectButton1.Location = New-Object System.Drawing.Size(280,18)
$SelectButton1.Size = New-Object System.Drawing.Size(75,23)
$SelectButton1.Text = "Select"
$SelectButton1.Add_Click({$objTextBox1.Text = Select-FileDialog})
$winform.Controls.Add($SelectButton1)


function RemoveAlias 
 {
 $global:fullpath = $objTextBox1.Text
 $global:filename = [System.IO.Path]::GetFileNameWithoutExtension("$fullpath")
 Remove-Item "HKLM:\software\microsoft\windows\currentversion\app paths\$filename.exe"
 if (Test-Path "HKLM:\software\microsoft\windows\currentversion\app paths\$filename.exe")
 {

 [Windows.Forms.MessageBox]::Show("Shortcut Not Removed", "Failed",  
       [Windows.Forms.MessageBoxButtons]::OK ,  
       [Windows.Forms.MessageBoxIcon]::Warning) 

}
 else
 {
 [Windows.Forms.MessageBox]::Show("Shortcut Removed", "Success",  
       [Windows.Forms.MessageBoxButtons]::OK ,  
       [Windows.Forms.MessageBoxIcon]::Information) 
 }
 }


$RemoveButton = New-Object System.Windows.Forms.Button
$RemoveButton.Location = New-Object System.Drawing.Size(150,110)
$RemoveButton.Size = New-Object System.Drawing.size(100,23)
$RemoveButton.Text = "Remove Alias"
$RemoveButton.add_click({RemoveAlias})
$WinForm.Controls.Add($RemoveButton)

function AddAlias
{
$global:fullpath = $objTextBox1.Text
$global:filename = [System.IO.Path]::GetFileNameWithoutExtension("$fullpath")

cd "HKLM:\software\microsoft\windows\currentversion\app paths\"
New-Item -Path "." -Name "$filename.exe" -Value "Default Value" -Force
Set-ItemProperty -Path ".\$filename.exe" -Name "(Default)" -Value "$fullpath"
#New-ItemProperty -Path ".\$filename.exe" -Name "$filename.exe" -Value "$fullpath"

if (Test-Path "HKLM:\software\microsoft\windows\currentversion\app paths\$filename.exe") 
{
[Windows.Forms.MessageBox]::Show("Shortcut created", "Success",  
       [Windows.Forms.MessageBoxButtons]::OK ,  
       [Windows.Forms.MessageBoxIcon]::Information)  
}
else {
[Windows.Forms.MessageBox]::Show("Failed to create Shortcut", "Failed",  
       [Windows.Forms.MessageBoxButtons]::OK ,  
       [Windows.Forms.MessageBoxIcon]::Warning) 
}
}

$AddAlias = New-Object System.Windows.Forms.Button
$AddAlias.Location = New-Object System.Drawing.Size(50,110)
$AddAlias.Size = New-Object System.Drawing.Size(75,23)
$AddAlias.Text = "Add Alias"
$AddAlias.Add_Click({AddAlias})
$winform.Controls.Add($AddAlias)

$CloseButton = New-Object System.Windows.Forms.Button
$CloseButton.Location = New-Object System.Drawing.Size(280,110)
$CloseButton.Size = New-Object System.Drawing.Size(75,23)
$CloseButton.Text = "Close"
$CloseButton.Add_Click({$WinForm.Close()})
$winform.Controls.Add($CloseButton)

$WinForm.Add_Shown($WinForm.Activate())  
$WinForm.showdialog() | out-null  
