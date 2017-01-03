##############################################################
# PowerShell user Script                  ver 0.2 25/01/2016 #
# Powershell Version 5.0                       Nathan Kewley #
#                                                            #
# Find and reports all users that are missing the manager    #
# and or title field. Return html response as well as log    #
##############################################################

function global:getMissingManagerOrTitle(){
    #enable cmdlet binding
    [CmdletBinding()]

    #define the parameters that the cmdlet will take
    param(
        [parameter(HelpMessage='true if automatic fixes should be actioned (Custom Common Parameter)')]
        [STRING]$action="false"
    )

    #create and start a timer for the script
    $timer = [Diagnostics.Stopwatch]::StartNew()

    #create var to track issues
    $global:issues = 0

    #default heading font for html report scripts
    $global:emailBody = "<font face=Arial Rounded MT Bold color=#ff6633 size=5>Missing Manager and Title Report</font><br><br>"

    #logfile for script
    $CurrentDate = Get-Date
    $CurrentDate = $CurrentDate.ToString('MM-dd-yyyy_hh-mm')
    $log = "C:\Logs\AD_MissingManagerOrTitle\Log_$CurrentDate.csv"

    #get all users with no manager or title field
    $users = Get-ADUser -LDAPFilter "(|(!manager=*)(!title=*))" -SearchScope OneLevel -SearchBase 'ou=user, ou=accounts, ou=new, dc=penrith, dc=local' -Properties UserPrincipalName, manager, title

    #check each user for title and manager
    foreach($user in $users){
        #check title
        if(-not($user.title -eq "*")){
            $logString = $user.UserPrincipalName + " Does not have a title"
            $logString >> $log
            $global:emailBody += $logString + "<br>"
            $global:issues++
        }

        #check manager
        if(-not($user.manager -eq "*")){
            $logString = $user.UserPrincipalName + " Does not have a manager"
            $logString >> $log
            $global:emailBody += $logString + "<br>"        }
            $global:issues++ 
    }

    #finalise html report
    $global:emailBody += "<font color=#686868 face=Arial size=3><br>$global:issues issues need to be resolved <br> "

    #write out final stuff
    $timer.Stop()
    $logString = "Script executed in " + $timer.Elapsed + "</font>"
    $global:emailBody += $logString

    #retuen the result as a html page
    return $global:emailBody
}