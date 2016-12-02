#Mindflash api gui v0.1 by nathan Kewley

#include my mindflash api backend
Import-Module $PSScriptRoot\get-MindflashUser.ps1

#ask the user for some form of credentials
$cred = Get-Credential

#create a glboal for the filename of the csv
$global:filename = "null"
$global:logfile = $PSScriptRoot + "\mindflashLog.txt"

#load form assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null

#Show an Open Folder Dialog and return the directory selected by the user.
function FolderBrowserDialog([string]$Message, [string]$InitialDirectory)
{
    $openFileDialog = New-Object System.Windows.Forms.openFileDialog
    $openFileDialog.ShowHelp = $true
    $openFileDialog.Title = "Select CSV"
    $OpenFileDialog.filter = "CSV Files (.csv) | *.csv"
    $openFileDialog.ShowDialog() | Out-Null
    $global:filename = $openFileDialog.FileName
    $form.Text = "$global:filename"
}

#funcition for single password reset
function resetSinglePasswd(){
    #set random if required
    if($passwd.text -eq "random"){
        $random = Get-Random -Minimum -1000 -Maximum 9999
        $passwd.text = "Penrith" + $random + "!"
    }

    #Make sure key is set
    if($api.text -eq "API KEY"){
        $output.appendText("ERROR: No API Key Supplied`r`n")
        return
    }
    
    #make sure password is set
    if(-not($passwd.text -ne "Enter New Password Here (type 'random' for random password)")){
        #error stuff out
        $output.appendText("ERROR: No Password Supplied`r`n")
        return
    }

    #make sure email is set
    if(-not($passwdEmail.text -ne "Enter User Email Here")){
        #error stuff out
        $output.appendText("ERROR: No Email Address Supplied`r`n")
        return
    }

    #try change the password
    else{
        try{
            #change the password
            get-MindflashUser -email $passwdEmail.text -resetPassword true -newPassword $passwd.text -key $api.text -resetAD false

            #send the email if required
            if ($check.Checked){

                #create message values
                $subject = "Connect Learn one-time computer generated unique password"
                $username = $cred.UserName

                #create the email
                $body = "<img src='logo.jpg'><br><br>Your new computer generated unique password to access Connect Learn is <strong>" + $passwd.text + "</strong><br><br>

                You can login to Connect Learn using your email address and the above password. If you have already accessed Connect Learn in the past and have already created a new password it will have been reset to the new unique password above.<br><br>

                This computer generated unique password <strong>must only be used once</strong> for the initial login. <br><br>

                After logging in with your unique password <strong>you must then create a new password</strong> under the edit profile link from the Trainee Dashboard. You can then use this new password you have created for any future login to Connect Learn.<br><br>

                For further instructions on how to access Connect Learn and how to use this computer generated unique password please refer to the 'Accessing Connect Learn - Childcare Centre v2c' instruction document sent to you in a separate email.<br><br>

                If you have any questions please contact the ICT Service Desk.<br><br>

                Thanks,<br><br>

                Harold Dulay<br>
                On behalf of the Connect Learn Project Team
                "

                #imagePath
                $imagePath = $PSScriptRoot + "\logo.jpg"

                #send the email
                send-mailmessage -to $passwdEmail.text -from $username -subject $subject -BodyAsHtml $body -smtpServer smtp.office365.com -Attachments $imagePath -credential $cred -useSSL  

                #tell user that email sent
                $email = "Email sent to: " + $passwdEmail.text + "`r`n"
                $output.appendText($email)
            }

            $output.appendText("Password for "+$passwdEmail.text+" Has been reset to "+$passwd.text+"`r`n")
        }catch{
            $output.appendText("Password Reset Faild: "+$passwdEmail.text+" to: "+$passwd.text+" EX: " + $_.Exception.Message + "`r`n")
        }
    }
}

#enroll single user
function enrollSingleUser(){
    #Make sure key is set
    if($api.text -eq "API KEY"){
        $output.appendText("ERROR: No API Key Supplied`r`n")
        return
    }
    
    #make sure course is set
    if(-not($singleCourse.text -ne "Enter Course Number Here")){
        #error stuff out
        $output.appendText("ERROR: No course number supplied`r`n")
        return
    }

    #make sure email is set
    if(-not($courseEmail.text -ne "Enter User Email Here")){
        #error stuff out
        $output.appendText("ERROR: No Email Address Supplied`r`n")
        return
    }

    #try enroll the user
    else{
        try{
            get-MindflashUser -email $courseEmail.text -course $singleCourse.text -invite true -key $api.text
            $output.appendText("User "+$courseEmail.text+" Has been enrolled in "+$singleCourse.text+"`r`n")
        }catch{
            $output.appendText("Could not enroll user: "+$courseEmail.text+" into: "+$singleCourse.text+"`r`n")
        }
    }
}

#create a single user
function createSingleUser(){
    #set random passwd if required
    if($passwd.text -eq "random"){
        $random = Get-Random -Minimum -1000 -Maximum 9999
        $passwd.text = "Penrith" + $random + "!"
    }

    #Make sure key is set
    if($api.text -eq "API KEY"){
        $output.appendText("ERROR: No API Key Supplied`r`n")
        return
    }
    
    #make sure first name is set
    if(-not($usrFirst.text -ne "First Name")){
        #error stuff out
        $output.appendText("ERROR: No first name supplied`r`n")
        return
    }

    #make sure last name is set
    if(-not($usrLast.text -ne "Last Name")){
        #error stuff out
        $output.appendText("ERROR: No Last Name Supplied`r`n")
        return
    }

    #make new password is set
    if(-not($usr.text -ne "Enter New Password Here (type 'random' for random password)")){
        #error stuff out
        $output.appendText("ERROR: No password supplied`r`n")
        return
    }

    #make sure email is set
    if(-not($usrEmail.text -ne "Enter User Email Here")){
        #error stuff out
        $output.appendText("ERROR: No Email Address Supplied`r`n")
        return
    }

    #try create the user
    else{
        try{
            get-MindflashUser -email $usrEmail.text -first $usrFirst.text -last $usrLast.text -newUser true -newPassword $usr.text -key $api.text
            $output.appendText("Succesfully created user: "+$usrEmail.text+"`r`n")
        }catch{
            $output.appendText("Could not create user: "+$usrEmail.text+"`r`n")
        }
    }
}

#function to reset bulk password
function resetBulkPasswd(){
    #create fail/pass variables
    $fail = 0
    $pass = 0

    #Make sure key is set
    if($api.text -eq "API KEY"){
        $output.appendText("ERROR: No API Key Supplied`r`n")
        return
    }
    
    #Make sure csv is selected
    if(($global:filename -eq "null") -or ($global:filename -eq "")){
        $output.appendText("ERROR: No CSV Selected`r`n")
        return
    }

    #make sure password is set
    if(-not($passBulk.text -ne "Enter New Password Here (type 'random' for random password)")){
        #error stuff out
        $output.appendText("ERROR: No Password Supplied`r`n")
        return
    }

    #open the selected csv file
    $users = Import-Csv $global:filename
    
    #cycle through each user in the csv
    Foreach ($user in $users){
        #set random password if required
        if($passBulk.text -eq "random"){
            $random = Get-Random -Minimum -1000 -Maximum 9999
            $newPass = "Penrith" + $random + "!"
        }else{$newPass = $passBulk.text}

        #build output strings
        $out = "user " + $user.email + " password has been reset to " + $newPass + "`r`n"
        $error = "ERROR: user " + $user.email + " password could not be reset`r`n"

        #try reset the users passwords
        try{
            #reset users password
            get-MindflashUser -email $user.email -resetPassword true -newPassword $newPass -key $api.text -resetAD false

            #only send out the email if the email box is ticked once we create a check box for this
            if ($check.Checked){

                #create message values
                $subject = "Connect Learn one-time computer generated unique password"
                $username = $cred.UserName

                #create the email
                $body = "<img src='logo.jpg'><br><br>Your new computer generated unique password to access Connect Learn is <strong>" + $newPass + "</strong>.<br><br>

                You can login to Connect Learn using your email address and the above password. If you have already accessed Connect Learn in the past and have already created a new password it will have been reset to the new unique password above.<br><br>

                This computer generated unique password <strong>must only be used once</strong> for the initial login. <br><br>

                After logging in with your unique password <strong>you must then create a new password</strong> under the edit profile link from the Trainee Dashboard. You can then use this new password you have created for any future login to Connect Learn.<br><br>

                For further instructions on how to access Connect Learn and how to use this computer generated unique password please refer to the 'Accessing Connect Learn - Childcare Centre v2c' instruction document sent to you in a separate email.<br><br>

                If you have any questions please contact the ICT Service Desk.<br><br>

                Thanks,<br><br>

                Harold Dulay<br>
                On behalf of the Connect Learn Project Team
                "

                #imagePath
                $imagePath = $PSScriptRoot + "\logo.jpg"

                #send the email
                send-mailmessage -to $user.email -from $username -subject $subject -BodyAsHtml $body -smtpServer smtp.office365.com -Attachments $imagePath -credential $cred -useSSL  

                #tell user that email sent
                $email = "Email sent to: " + $user.email + "`r`n"
                $output.appendText($email)
            }

            $output.appendText($out)
            $pass++;
        }catch{
            $output.appendText($error)
            $fail++;
        }
    }

    #output pass/fail rate
    $output.AppendText("Total Success: $pass | Total Fail: $fail`r`n")
}

#function to buld enroll users
function bulkEnroll(){
    #create fail/pass variables
    $fail = 0
    $pass = 0

    #Make sure key is set
    if($api.text -eq "API KEY"){
        $output.appendText("ERROR: No API Key Supplied`r`n")
        return
    }
    
    #Make sure csv is selected
    if(($global:filename -eq "null") -or ($global:filename -eq "")){
        $output.appendText("ERROR: No CSV Selected`r`n")
        return
    }

    #make sure course is set
    if(-not($course.text -ne "Enter Course Number Here")){
        #error stuff out
        $output.appendText("ERROR: No Course Code Supplied Supplied`r`n")
        return
    }

    #open the selected csv file
    $users = Import-Csv $global:filename
    
    #cycle through each user in the csv
    Foreach ($user in $users){
        #build output strings
        $out = "user " + $user.email + " Has been enrolled in " + $course.text + "`r`n"
        $error = "ERROR: user " + $user.email + " could not be enrolled in course " + $course.text + "`r`n"

        #try to enroll the user
        try{
            get-MindflashUser -email $user.email -course $course.text -invite true -key $api.text
            $output.AppendText($out)
            $pass++;
        }catch{
            $output.AppendText($error)
            $fail++;
        }
    }

    #output pass/fail rate
    $output.AppendText("Total Success: $pass | Total Fail: $fail`r`n")
}

#function to bulk create users
function bulkCreate(){
    #create fail/pass variables
    $fail = 0
    $pass = 0

    #Make sure key is set
    if($api.text -eq "API KEY"){
        $output.appendText("ERROR: No API Key Supplied`r`n")
        return
    }
    
    #Make sure csv is selected
    if(($global:filename -eq "null") -or ($global:filename -eq "")){
        $output.appendText("ERROR: No CSV Selected`r`n")
        return
    }

    #Make sure key is set
    if($usrBulk.text -eq "Enter New Password Here (type 'random' for random password)"){
        $output.appendText("ERROR: No Password Supplied`r`n")
        return
    }

    #open the selected csv file
    $users = Import-Csv $global:filename
    
    #cycle through each user in the csv
    Foreach ($user in $users){
        #set random password if required
        if($usrBulk.text -eq "random"){
            $random = Get-Random -Minimum -1000 -Maximum 9999
            $newPass = "Penrith" + $random + "!"
        }else{$newPass = $usrBulk.text}

        #build output strings
        $out = "user " + $user.email + " Has been created" + "`r`n"
        $error = "ERROR: user " + $user.email + " Could not be created`r`n"

        #try reset the users passwords
        try{
            get-MindflashUser -email $user.email -first $user.first -last $user.last -newUser true -newPassword $newPass -key $api.text
            $output.appendText($out)
            $pass++;
        }catch{
            $output.appendText($error)
            $fail++;
        }
    }

    #output pass/fail rate
    $output.AppendText("Total Success: $pass | Total Fail: $fail`r`n")

}

####################-------GUI STUFF BELOW--------####################

#declare font and icon
$Font = New-Object System.Drawing.Font("Times New Roman",10,[System.Drawing.FontStyle]::Bold)
$Icon = New-Object system.drawing.icon ("$PSScriptRoot\icon.ico")

#create the form
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object Windows.Forms.Form
$form.Size = New-Object Drawing.Size @(640,480)
$Form.Icon = $Icon 
$form.StartPosition = "CenterScreen"
$form.Text = "One GUI To Rule Them All"

#create the select csv button
$csv = New-Object System.Windows.Forms.Button
$csv.Size = New-Object Drawing.Size @(100,20)
$csv.Location = New-Object System.Drawing.Size(12,5) 
$csv.add_click({FolderBrowserDialog})
$csv.Text = "Select CSV"

#create api key box
$api = New-Object System.Windows.Forms.TextBox 
$api.Location = New-Object System.Drawing.Size(124,5) 
$api.Size = New-Object System.Drawing.Size(488,20) 
$api.Text = "API KEY"

#create enroll tab stuff
$course = New-Object System.Windows.Forms.TextBox 
$course.Location = New-Object System.Drawing.Size(13,25) 
$course.Size = New-Object System.Drawing.Size(387,20) 
$course.Text = "Enter Course Number Here"

$singleCourse = New-Object System.Windows.Forms.TextBox 
$singleCourse.Location = New-Object System.Drawing.Size(13,95) 
$singleCourse.Size = New-Object System.Drawing.Size(387,20) 
$singleCourse.Text = "Enter Course Number Here"

$courseEmail = New-Object System.Windows.Forms.TextBox 
$courseEmail.Location = New-Object System.Drawing.Size(13,70) 
$courseEmail.Size = New-Object System.Drawing.Size(563,20) 
$courseEmail.Text = "Enter User Email Here"

$courseBTN = New-Object System.Windows.Forms.Button
$courseBTN.Size = New-Object Drawing.Size @(164,20)
$courseBTN.Location = New-Object System.Drawing.Size(413,25) 
$courseBTN.add_click({bulkEnroll})
$courseBTN.Text = "Bulk Enroll Users"

$singleCourseBTN = New-Object System.Windows.Forms.Button
$singleCourseBTN.Size = New-Object Drawing.Size @(164,20)
$singleCourseBTN.Location = New-Object System.Drawing.Size(413,95) 
$singleCourseBTN.add_click({enrollSingleUser})
$singleCourseBTN.Text = "Enroll Single User"

$singleEnroll = New-Object System.Windows.Forms.Label
$singleEnroll.Location = New-Object System.Drawing.Size(237,50) 
$singleEnroll.Size = New-Object System.Drawing.Size(400,20) 
$singleEnroll.Font = $Font
$singleEnroll.Text = "Enroll A Single User"

$bulkEnroll = New-Object System.Windows.Forms.Label
$bulkEnroll.Location = New-Object System.Drawing.Size(187,5) 
$bulkEnroll.Size = New-Object System.Drawing.Size(400,20) 
$bulkEnroll.Font = $Font
$bulkEnroll.Text = "Bulk Enroll Users From selected CSV"

#create passwd reset stuff
$checkLabel = New-Object System.Windows.Forms.Label
$checkLabel.Location = New-Object System.Drawing.Size(13,125) 
$checkLabel.Size = New-Object System.Drawing.Size(400,20) 
$checkLabel.Text = "Send Email: "

$check = New-Object System.Windows.Forms.CheckBox
$check.Location = New-Object System.Drawing.Size(77,120) 

$passwd = New-Object System.Windows.Forms.TextBox 
$passwd.Location = New-Object System.Drawing.Size(13,95) 
$passwd.Size = New-Object System.Drawing.Size(387,20) 
$passwd.Text = "Enter New Password Here (type 'random' for random password)"

$passBulk = New-Object System.Windows.Forms.TextBox 
$passBulk.Location = New-Object System.Drawing.Size(13,25) 
$passBulk.Size = New-Object System.Drawing.Size(387,20) 
$passBulk.Text = "Enter New Password Here (type 'random' for random password)"

$passwdEmail = New-Object System.Windows.Forms.TextBox 
$passwdEmail.Location = New-Object System.Drawing.Size(13,70) 
$passwdEmail.Size = New-Object System.Drawing.Size(563,20) 
$passwdEmail.Text = "Enter User Email Here"

$passwdBTN = New-Object System.Windows.Forms.Button
$passwdBTN.Size = New-Object Drawing.Size @(164,20)
$passwdBTN.Location = New-Object System.Drawing.Size(413,25) 
$passwdBTN.add_click({resetBulkPasswd})
$passwdBTN.Text = "Bulk Reset Passwords"

$singlePasswdBTN = New-Object System.Windows.Forms.Button
$singlePasswdBTN.Size = New-Object Drawing.Size @(164,20)
$singlePasswdBTN.Location = New-Object System.Drawing.Size(413,95) 
$singlePasswdBTN.add_click({resetSinglePasswd})
$singlePasswdBTN.Text = "Reset Single Password"

$singlePasswd = New-Object System.Windows.Forms.Label
$singlePasswd.Location = New-Object System.Drawing.Size(210,50) 
$singlePasswd.Size = New-Object System.Drawing.Size(400,20) 
$singlePasswd.Font = $Font
$singlePasswd.Text = "Reset A Single Users Password"

$bulkPasswd = New-Object System.Windows.Forms.Label
$bulkPasswd.Location = New-Object System.Drawing.Size(195,5) 
$bulkPasswd.Size = New-Object System.Drawing.Size(400,20) 
$bulkPasswd.Font = $Font
$bulkPasswd.Text = "Bulk Reset Passwords From A CSV"

#create create userstuff
$usr = New-Object System.Windows.Forms.TextBox 
$usr.Location = New-Object System.Drawing.Size(13,120) 
$usr.Size = New-Object System.Drawing.Size(387,20) 
$usr.Text = "Enter New Password Here (type 'random' for random password)"

$usrBulk = New-Object System.Windows.Forms.TextBox 
$usrBulk.Location = New-Object System.Drawing.Size(13,25) 
$usrBulk.Size = New-Object System.Drawing.Size(387,20) 
$usrBulk.Text = "Enter New Password Here (type 'random' for random password)"

$usrEmail = New-Object System.Windows.Forms.TextBox 
$usrEmail.Location = New-Object System.Drawing.Size(13,95) 
$usrEmail.Size = New-Object System.Drawing.Size(563,20) 
$usrEmail.Text = "Enter User Email Here"

$usrFirst = New-Object System.Windows.Forms.TextBox 
$usrFirst.Location = New-Object System.Drawing.Size(13,70) 
$usrFirst.Size = New-Object System.Drawing.Size(277,20) 
$usrFirst.Text = "First Name"

$usrLast = New-Object System.Windows.Forms.TextBox 
$usrLast.Location = New-Object System.Drawing.Size(300,70) 
$usrLast.Size = New-Object System.Drawing.Size(276,20) 
$usrLast.Text = "Last Name"

$usrBTN = New-Object System.Windows.Forms.Button
$usrBTN.Size = New-Object Drawing.Size @(164,20)
$usrBTN.Location = New-Object System.Drawing.Size(413,25) 
$usrBTN.add_click({bulkCreate})
$usrBTN.Text = "Bulk Create Users"

$singleUsrBTN = New-Object System.Windows.Forms.Button
$singleUsrBTN.Size = New-Object Drawing.Size @(164,20)
$singleUsrBTN.Location = New-Object System.Drawing.Size(413,120) 
$singleUsrBTN.add_click({createSingleUser})
$singleUsrBTN.Text = "Create Single User"

$singleUsr = New-Object System.Windows.Forms.Label
$singleUsr.Location = New-Object System.Drawing.Size(220,50) 
$singleUsr.Size = New-Object System.Drawing.Size(400,20) 
$singleUsr.Font = $Font
$singleUsr.Text = "Create A Single User"

$bulkUsr = New-Object System.Windows.Forms.Label
$bulkUsr.Location = New-Object System.Drawing.Size(195,5) 
$bulkUsr.Size = New-Object System.Drawing.Size(400,20) 
$bulkUsr.Font = $Font
$bulkUsr.Text = "Bulk Create Users From A CSV"

#create enroll tab
$tabEnroll = New-Object System.Windows.Forms.TabPage
$tabEnroll.UseVisualStyleBackColor = $True
$tabEnroll.Text = "Enroll Users”
$tabEnroll.Controls.Add($course)
$tabEnroll.Controls.Add($courseBTN)
$tabEnroll.Controls.Add($bulkEnroll)
$tabEnroll.Controls.Add($singleEnroll)
$tabEnroll.Controls.Add($courseEmail)
$tabEnroll.Controls.Add($singleCourse)
$tabEnroll.Controls.Add($singleCourseBTN)

#create reset passwd tab
$tabPasswd = New-Object System.Windows.Forms.TabPage
$tabPasswd.UseVisualStyleBackColor = $True
$tabPasswd.Text = "Reset Passwords”
$tabPasswd.Controls.Add($bulkPasswd)
$tabPasswd.Controls.Add($passBulk)
$tabPasswd.Controls.Add($passwdBTN)
$tabPasswd.Controls.Add($passwdEmail)
$tabPasswd.Controls.Add($passwd)
$tabPasswd.Controls.Add($singlePasswd)
$tabPasswd.Controls.Add($singlePasswdBTN)
$tabPasswd.Controls.Add($check)
$tabPasswd.Controls.Add($checkLabel)

#create create users tab
$tabUser = New-Object System.Windows.Forms.TabPage
$tabUser.UseVisualStyleBackColor = $True
$tabUser.Text = "Create Users”
$tabUser.Controls.Add($usrBulk)
$tabUser.Controls.Add($usrBTN)
$tabUser.Controls.Add($bulkUsr)
$tabUser.Controls.Add($usrFirst)
$tabUser.Controls.Add($usrLast)
$tabUser.Controls.Add($singleUsr)
$tabUser.Controls.Add($usrEmail)
$tabUser.Controls.Add($usr)
$tabUser.Controls.Add($singleUsrBTN)

#create the tab form
$tab = New-Object System.Windows.Forms.TabControl
$tab.Location = New-Object System.Drawing.Size(12,35) 
$tab.Size = New-Object System.Drawing.Size(600,180) 
$tab.Controls.Add($tabPasswd)
$tab.Controls.Add($tabEnroll)
$tab.Controls.Add($tabUser)

#create output text box
$output = New-Object System.Windows.Forms.TextBox 
$output.Location = New-Object System.Drawing.Size(12,220) 
$output.Multiline = $True
$output.ScrollBars = "Vertical"
$output.Size = New-Object System.Drawing.Size(599,215) 
$output.text = "Mindflash API GUI v0.1 | Nathan Kewley | 17-12-2015`r`n"

#add elements and show the form
$form.Controls.Add($csv) 
$form.Controls.Add($api) 
$form.Controls.Add($tab) 
$form.Controls.Add($output) 
$drc = $form.ShowDialog()