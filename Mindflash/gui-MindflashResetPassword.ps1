#include my mindfla`r`nsh api backend
Import-Module $PSScriptRoot\get-MindflashUser.ps1

#hardcode credentials like a fearless idiot
#$username = "nathan.kewley@penrith.city"
#$password = "Stripe175"
#$secstr = New-Object -TypeName System.Security.SecureString
#$password.ToCharArray() | ForEach-Object {$secstr.AppendChar($_)}
#$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $secstr

$cred = Get-Credential

#load form assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

#dirty globals
$pass = 0
$fail = 0
$ad = "false"
$rand = "false"

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

    #see if the ad password should also be reset
    if($passDrop.selectedItem.ToString() -eq "Mindflash + AD"){
        $ad = "true"
    }elseif($passDrop.selectedItem.ToString() -eq "AD Only"){
        $ad = "only"
    }

    if($password.text -eq "random"){
        $rand = "true"
    }else{
        $rand = "false"
    }

    #check if adding by email or sam
    if($drop.SelectedItem.ToString() -eq "SAM"){
        #for each staff member read from the csv file
        Foreach ($user in $users){
            #set random password if required
            if($rand -eq "true"){
                $random = Get-Random -Minimum -1000 -Maximum 9999
                $password.text = "Penrith" + $random + "!"
            }

            #create success output string
            $out = "user " + $user.sam + " password has been reset to " + $password.text + "`r`n"

            #create error string
            $error = "ERROR: user " + $user.sam + " password could not be reset`r`n"

            #try add the users giving output as to wether it was a success
            try{
                get-MindflashUser -sam $user.sam -resetPassword true -newPassword $password.text -key $api.text -resetAD $ad

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
    }elseif($drop.SelectedItem.ToString() -eq "EMAIL"){
        #for each staff member read from the csv file
        Foreach ($user in $users){
            #set random password if required
            if($rand -eq "true"){
                $random = Get-Random -Minimum 1000 -Maximum 9999
                $password.text = "Penrith" + $random + "!"
            }

            #create success output string
            $out = "user " + $user.email + " password has been reset`r`n" + $password.text + "`r`n"

            #create error string
            $error = "ERROR: user " + $user.email + " password could not be reset`r`n"

            #try add the users giving output as to wether it was a success
            try{
                get-MindflashUser -email $user.email -resetPassword true -newPassword $password.text -key $api.text -resetAD $ad

                #Create the email
                $body = "Your new computer generated unique password to access Connect Learn is " + $password.text + "<br><br>.

You can login to Connect Learn using your email address and the above password. If you have already accessed Connect Learn in the past and have already created a new password it will have been reset to the new unique password above.<br><br>

This computer generated unique password <strong>must only be used once</strong> for the initial login. <br><br>

After logging in with your unique password <strong>you must then create a new password</strong> under the edit profile link from the Trainee Dashboard. You can then use this new password you have created for any future login to Connect Learn.<br><br>

For further instructions on how to access Connect Learn and how to use this computer generated unique password please refer to the 'Accessing Connect Learn - Childcare Centre v2c' instruction document sent to you in a separate email.<br><br>

If you have any questions please contact the ICT Service Desk.<br><br>

Thanks,<br><br>

Harold Dulay<br>
Connect Learn Team
"

                $subject = "Connect Learn one-time computer generated unique password"

                #send the email
                send-mailmessage -to $user.email -from $username -subject $subject -BodyAsHtml $body -smtpServer smtp.office365.com -credential $cred -useSSL  

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
$form.Text = "csv: $filename | Reset Password"

#create password text box
$password = New-Object System.Windows.Forms.TextBox 
$password.Location = New-Object System.Drawing.Size(330,5) 
$password.Size = New-Object System.Drawing.Size(195,20) 
$password.Text = "Enter New Password Here"

#create api key box
$api = New-Object System.Windows.Forms.TextBox 
$api.Location = New-Object System.Drawing.Size(13,35) 
$api.Size = New-Object System.Drawing.Size(513,20) 
$api.Text = "Enter API key here"

#create start button
$btn = New-Object System.Windows.Forms.Button
$btn.Size = New-Object Drawing.Size @(100,20)
$btn.Location = New-Object System.Drawing.Size(12,5) 
$btn.add_click({beAMindflashGod})
$btn.Text = "Reset Passwords"

#create email/sam dropdown box
$drop = New-Object System.Windows.Forms.ComboBox
$drop.Location = New-Object System.Drawing.Size(125,5)
$drop.Size = New-Object System.Drawing.Size(75,20)
$drop.DropDownHeight = 200
#$drop.Items.Add("SAM") only allow from email for now
$drop.Items.Add("EMAIL")
$drop.selectedIndex = 0

#create dropdown for which passwords to reset
$passDrop = New-Object System.Windows.Forms.ComboBox
$passDrop.Location = New-Object System.Drawing.Size(213,5)
$passDrop.Size = New-Object System.Drawing.Size(110,20)
$passDrop.DropDownHeight = 200
$passDrop.Items.Add("Mindflash Only")
$passDrop.Items.Add("AD Only")
$passDrop.Items.Add("Mindflash + AD")
$passDrop.selectedIndex = 0

#create output text box
$output = New-Object System.Windows.Forms.TextBox 
$output.Location = New-Object System.Drawing.Size(12,70) 
$output.Multiline = $True
$output.ScrollBars = "Vertical"
$output.Size = New-Object System.Drawing.Size(512,150) 
$output.text = "Note: that reset emails will only be sent out if the email option is selected`r`n"

#add elements and show the form
$form.Controls.Add($password) 
$form.Controls.Add($btn)
$form.Controls.Add($output) 
$form.Controls.Add($api) 
$Form.Controls.Add($drop) 
$form.Controls.Add($passDrop) 
$drc = $form.ShowDialog()