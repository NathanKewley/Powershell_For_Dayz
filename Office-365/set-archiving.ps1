##############################################################
# Enable Archiving                         ver 0.1 08/04/2016#
# Powershell Version 4.0                                     #
# If this script works it was written by Nathan Kewley       #
# If this script does not work is was written by Michael Cho #
#                                                            #
# get all users that do not have archiving  enabled and      #
# generate a report of these users, then enable archiving    #
# for these users.                                           #
#                                                            #
# you will need to be connected to 365 for this to work      #
##############################################################

function enable-archive(){
    #some counter vars
    $issues = 0
    $enabled = 0
    $timer = [Diagnostics.Stopwatch]::StartNew()

    #get credentials from the user for email sending
    $cred = Get-Credential
    
    #get all users who do not have archiving enabled
    $mailboxes = Get-Mailbox -ResultSize 5 -Filter {RecipientTypeDetails -eq "UserMailbox"} | Where {$_.archivestatus -match "none"} | ft

    #create the html body for the email
    $emailBody = "<font face=Arial Rounded MT Bold color=#ff6633 size=6>Mailbox Archive Report</font><hr><br>"
    $emailBody += "<font face=Arial Rounded MT Bold color=#ff6633 size=4>USERS WITHOUT MAILBOX ARCHIVING</font><br>"

    #Interate through all users who do not have legal hold enables and add them to the report
    foreach($box in $mailboxes){
        #try to enable archiving for the user
        try{
            #enable archiving
            Enable-Mailbox "$box" -Archive -WhatIf

            #add success to the report
            $emailBody += "<br><font color=#339933>$box - ENABLED</font>"
            $enabled++
        }

        #if archive could not be set for user we should error out
        catch{
            #add fail to the report
            $emailBody += "<br><font color=#800000>$box - COULD NOT ENABLE ARCHIVING</font>"
            $issues++
        }
    }

    #add info to the bottom of the script
    $emailBody += "<br><br> $enabled Users Have Been Enabled"
    if($issues > 0){$emailBody += "<br><font color=#800000>$issues Users Could Not Be Enabled</font><br>"}
    else{$emailBody += "<br>$issues Users Could Not Be Enabled<br>"}
    $timer.Stop()
    $emailBody += "Script executed in " + $timer.Elapsed + " seconds"

    #Send the report email
    $Recipients = "That Bawss <michael.cho@penrith.city>", "scripter <nathan.kewley@penrith.city>"
    send-mailmessage -to $Recipients -from "nathan.kewley@penrith.city" -subject "Mailbox Archive Report" -bodyashtml $emailBody -smtpServer smtp.office365.com -credential $cred -useSSL -priority High
}