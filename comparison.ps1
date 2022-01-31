<#

***************   This script is provided AS-IS without any warranty to any damage that may occured. It will delete emails, if you are using it it's AT YOUR OWN RISK!  ***********

Version 1.0
Inital release

#>


Function HR_Sel_File
{
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = 'CSV (*.csv)|*.csv'
    }
    If ($FileBrowser.ShowDialog() -eq "Cancel") 
    {
        [System.Windows.Forms.MessageBox]::Show("Please select the HR file !", "Error", 0, 
        [System.Windows.Forms.MessageBoxIcon]::Exclamation)
    }
    $Global:HR_SelectedFile = $FileBrowser.FileName
    $ChooseHR.width = 450

}

Function 365_Sel_File
{
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'CSV (*.csv)|*.csv'
    }
    If ($FileBrowser.ShowDialog() -eq "Cancel") 
    {
        [System.Windows.Forms.MessageBox]::Show("Please select the 365 file !", "Error", 0, 
        [System.Windows.Forms.MessageBoxIcon]::Exclamation)
    }
    $Global:365_SelectedFile = $FileBrowser.FileName
    $Choose365.width = 450
}

Function check()
{
    $365_Active_Users =@()
    $HR_ActiveUsers =@()
    $UserPath = "$($env:USERPROFILE)\Desktop\"

    #Import the HR file
    $HR_Users = Import-Csv $Global:HR_SelectedFile
    $HR_ActiveUsers = $HR_Users."Business Email"

    #Imprt the 365 file
    $365_Users = Import-Csv $Global:365_SelectedFile
    ForEach ($365_User in $365_Users)
    {
        #Check if the user signin is blocked
        if ($365_User.BlockCredential -Like "False" -and $365_User.PreferredDataLocation -notcontains '' -and $365_User.IsLicensed -Like "True")
        {
            $365_Active_Users += $365_User.SignInName
        }
    }
    #compare the rrays and export the results
    Compare-Object -ReferenceObject $365_Active_Users -DifferenceObject $HR_ActiveUsers | Where-Object {$_.SideIndicator -eq "<="} | Select-Object InputObject | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Set-Content $env:temp\NotInTW.csv
    [System.Windows.Forms.MessageBox]::Show('The comparison has been completed','365 Users report','Ok','Information')
    $objExcel = New-Object -ComObject Excel.Application
    $folder = $env:temp
    $file = "\NotInTW.csv"
    $path = join-path $folder $file
    $objExcel.Workbooks.Open($path)
    $objExcel.Visible = $true
    $Form.Close()
}

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.AutoSize                   = $true
$Form.text                       = "365 Users report"
$Form.TopMost                    = $true
#----------------------

$ChooseHR_H                      = New-Object system.Windows.Forms.Label
$ChooseHR_H.text                 = "HR Data"
$ChooseHR_H.AutoSize             = $true
$ChooseHR_H.width                = 25
$ChooseHR_H.height               = 10
$ChooseHR_H.location             = New-Object System.Drawing.Point(28,20)
$ChooseHR_H.ForeColor            = "#000000"

$ChooseHR                        = New-Object System.Windows.Forms.Button
$ChooseHR.text                   = "Select the HR file"
$ChooseHR.AutoSize               = $true
$ChooseHR.width                  = 90
$ChooseHR.height                 = 20
$ChooseHR.location               = New-Object System.Drawing.Point(80,17)


$ChooseHR.Add_Click({HR_Sel_File
$ChooseHR.Text = $Global:HR_SelectedFile
}) 

$Choose365_H                      = New-Object system.Windows.Forms.Label
$Choose365_H.text                 = "365 Data"
$Choose365_H.AutoSize             = $true
$Choose365_H.width                = 25
$Choose365_H.height               = 10
$Choose365_H.location             = New-Object System.Drawing.Point(28,85)
$Choose365_H.ForeColor            = "#000000"

$Choose365                       = New-Object System.Windows.Forms.Button
$Choose365.text                   = "Select the 365 File"
$Choose365.AutoSize               = $true
$Choose365.width                  = 90
$Choose365.height                 = 20
$Choose365.location               = New-Object System.Drawing.Point(80,82)

$Choose365.Add_Click({365_Sel_File
$Choose365.Text = $Global:365_SelectedFile
}) 

#----------
$Apply                         = New-Object system.Windows.Forms.Button
$Apply.text                    = "Search"
$Apply.width                   = 99
$Apply.height                  = 30
$Apply.location                = New-Object System.Drawing.Point(200,190)
$apply.Add_Click({check})

#----------
$Cancel                         = New-Object system.Windows.Forms.Button
$Cancel.text                    = "Close"
$Cancel.width                   = 98
$Cancel.height                  = 30
$Cancel.location                = New-Object System.Drawing.Point(330,190)
$Cancel.Add_Click({$Form.Close()})

$Form.Controls.AddRange(@($ChooseHR, $ChooseHR_H, $Apply, $Cancel, $Choose365_H, $Choose365))
[void] $Form.ShowDialog()



