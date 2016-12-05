##############################################################
# PowerShell 365 licence ver 0.1 08/07/2015                  #
# Powershell Version 4.0                                     #
# If this script works it was written by Nathan Kewley       #
# If this script does not work is was written by Michael Cho #
#                                                            #
# License all Asure AD users with 365                        #
##############################################################
#needs msol extention for powershell
#Before running this script in powershell you need to connect to MsolService. Execute the following
#command in powershell and authenticate, Then run this script
#Connect-MsolService

#define file to write users who have been modified into
$filename = "C:\Users\P102542\Desktop\userLicense.csv"

#Output the date and time of script execution to the file
Get-Date >> $filename

#Get all unlicensed users (up to a max of 2000 users)
$users = Get-MsolUser -MaxResults 2000

#loop thorugh each user
foreach ($user in $users) {
        
    #Set usage location. This is required for every individual user.
    Set-MsolUser -UserPrincipalName $user.UserPrincipalName -UsageLocation "AU";
            
    #Apply enterprise pack license to all users
    Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses penrithcitycouncil:EMS;

    #create changes string
    $output = "user: " + $user.UserPrincipalName + " Has been given the license: penrithcitycouncil:EMS";

    #write changes to host
    write-host($output);

    #output changes users
    $output >> $filename;
}

#Write console output
write-host "Execution complete" -foregroundcolor "red"