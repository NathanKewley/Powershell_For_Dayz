#include my mindfla`r`nsh api backend
Import-Module $PSScriptRoot\get-MindflashUser.ps1

#load form assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

#dirty globals
$pass = 0;
$fail = 0;

#function to be a mindflash god
function beAMindflashGod()
{
    #start a log of what is happening
    $logfile = $PSScriptRoot + "\mindflashLog.txt"
    #$logfile +=  + "\mindflashLog.txt"
    #Start-Transcript -Path $logfile
    "----------Start of Transcript-------------" >> $logfile

    #open the selected csv file
    $users = Import-Csv $filename

    #check if adding by email or sam
    if($drop.SelectedItem.ToString() -eq "SAM"){
        #for each staff member read from the csv file
        Foreach ($user in $users){
            #create success output string
            $out = "user " + $user.sam + " has been added to course " + $course.text + "`r`n"

            #create error string
            $error = "ERROR: user " + $user.sam + " could not be added to course " + $course.text + "`r`n"

            #try add the users giving output as to wether it was a success
            try{
                get-MindflashUser -sam $user.sam -course $course.text -invite true -key $api.text

                #write the output
                write-host $out
                $out >> $logfile
                $output.Text += $out            
                $pass++;
            }catch{
                $output.Text += $error            
                $error >> $logfile
                $fail++;
            }
        }        
    }elseif($drop.SelectedItem.ToString() -eq "EMAIL"){
        #for each staff member read from the csv file
        Foreach ($user in $users){
            #create success output string
            $out = "user " + $user.email + " has been added to course " + $course.text + "`r`n"

            #create error string
            $error = "ERROR: user " + $user.email + " could not be added to course " + $course.text + "`r`n"

            #try add the users giving output as to wether it was a success
            try{
                get-MindflashUser -email $user.email -course $course.text -invite true -key $api.text

                #write the output
                write-host $out
                $out >> $logfile
                $output.Text += $out            
                $pass++;
            }catch{
                $error >> $logfile
                $output.Text += $error            
                $fail++;
            }
        }        
    }

    #output pass/fail rate
    $output.text += "Total Success: $pass | Total Fail: $fail`r`n"

    #tell the user where the full output log can be found
    $output.Text += "Error log: $logfile`r`n"
}

#Show an Open Folder Dialog and return the directory selected by the user.
function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory)
{
    $openFileDialog = New-Object System.Windows.Forms.openFileDialog
    $openFileDialog.ShowHelp = $true
    $openFileDialog.Title = "Select CSV for mindflash users"
    $openFileDialog.ShowDialog() | Out-Null
    return $openFileDialog.FileName
}

#get harold to select a csv
$filename = Read-FolderBrowserDialog
write-host $filename

#create the form
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object Windows.Forms.Form
$form.Size = New-Object Drawing.Size @(550,265)
$form.StartPosition = "CenterScreen"
$form.Text = "csv: $filename | Enroll Course"

#create course text box
$course = New-Object System.Windows.Forms.TextBox 
$course.Location = New-Object System.Drawing.Size(213,5) 
$course.Size = New-Object System.Drawing.Size(312,20) 
$course.Text = "Enter Course Number Here"

#create api key box
$api = New-Object System.Windows.Forms.TextBox 
$api.Location = New-Object System.Drawing.Size(12,35) 
$api.Size = New-Object System.Drawing.Size(512,20) 
$api.Text = "Enter API key here"

#create start button
$btn = New-Object System.Windows.Forms.Button
$btn.Size = New-Object Drawing.Size @(100,20)
$btn.Location = New-Object System.Drawing.Size(12,5) 
$btn.add_click({beAMindflashGod})
$btn.Text = "Enroll Users"

#create email/sam dropdown box
$drop = New-Object System.Windows.Forms.ComboBox
$drop.Location = New-Object System.Drawing.Size(125,5)
$drop.Size = New-Object System.Drawing.Size(75,20)
$drop.DropDownHeight = 200
#$drop.Items.Add("SAM") only allow from email from the gui
$drop.Items.Add("EMAIL")
$drop.selectedIndex = 0

#create output text box
$output = New-Object System.Windows.Forms.TextBox 
$output.Location = New-Object System.Drawing.Size(12,70) 
$output.Multiline = $True
$output.ScrollBars = "Vertical"
$output.Size = New-Object System.Drawing.Size(512,150) 

#add elements and show the form
$form.Controls.Add($course) 
$form.Controls.Add($api) 
$form.Controls.Add($btn)
$form.Controls.Add($output) 
$Form.Controls.Add($drop) 
$drc = $form.ShowDialog()