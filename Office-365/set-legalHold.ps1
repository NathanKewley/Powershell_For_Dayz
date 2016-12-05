##############################################################
# Enable Legal Hold                        ver 0.1 08/04/2016#
# Powershell Version 4.0                                     #
# If this script works it was written by Nathan Kewley       #
# If this script does not work is was written by Michael Cho #
#                                                            #
# get all users that do not have legal hold enabled and      #
# generate a report of these users, then enable legal hold   #
# for these users.                                           #
#                                                            #
# you will need to be connected to 365 for this to work      #
##############################################################

function enable-legalHold(){
    #some counter vars
    $issues = 0
    $enabled = 0
    $timer = [Diagnostics.Stopwatch]::StartNew()

    #get credentials from the user for email sending
    $cred = Get-Credential
    
    #get all users who do not have legal hold enabled
    $mailboxes = Get-Mailbox -ResultSize 5 -Filter {RecipientTypeDetails -eq "UserMailbox"} | Where {$_.LitigationHoldEnabled -match "False"}

    #create the html body for the email
    $emailBody = "<font face=Arial Rounded MT Bold color=#ff6633 size=6>Legal Hold Report</font><hr><br>"
    $emailBody += "<font face=Arial Rounded MT Bold color=#ff6633 size=4>USERS WITHOUT LEGAL HOLD</font><br>"

    #Interate through all users who do not have legal hold enables and add them to the report
    foreach($box in $mailboxes){
        #try to enable legal hold for the user
        try{
            #set legal hold for 7.5 years
            Set-Mailbox "$box" -LitigationHoldEnabled $True -LitigationHoldDuration 2737 -whatif
            
            #add success to the report
            $emailBody += "<br><font color=#339933>$box - ENABLED</font>"
            $enabled++
        }

        #if legal hold could not be set for user we should error out
        catch{
            #add fail to the report
            $emailBody += "<br><font color=#800000>$box - COULD NOT ENABLE LEGAL HOLD</font>"
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
    send-mailmessage -to $Recipients -from "nathan.kewley@penrith.city" -subject "Legal Hold Report" -bodyashtml $emailBody -smtpServer smtp.office365.com -credential $cred -useSSL -priority High
}