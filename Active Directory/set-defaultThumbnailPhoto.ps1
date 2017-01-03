##############################################################
# PowerShell AD Photo Update Script Ver 0.1 01/07/2015       #
# Powershell Version 4.0                                     #
# If this script works it was written by Nathan Kewley       #
# If this script does not work is was written by Michael Cho #
#                                                            #
# This script detects users in an OU who do not currently    #
# have a profile photo in AD and uploads a default photo.    #
# Users who already have a photo will not be modified.       #
# the 'samaccountname' of all modified users will be output  #
# to a text file.                                            #
##############################################################

#define file to write users who have been modified into
$filename = "C:\Users\P102542\Desktop\ChangesUsers.txt";

#Output the date and time of script execution to the file
Get-Date >> $filename

#Read the picture from file
$Picture=[System.IO.File]::ReadAllBytes('C:\orangeGuy.jpg')

#Get each user in the test OU that does not have a thumbnail picture
$list=Get-ADuser -Filter * -SearchBase 'ou=AD Cleanup, ou=test, ou=accounts, ou=new, dc=penrith, dc=local' -properties thumbnailPhoto | ? {!$_.thumbnailPhoto}

#For each user in the OU
Foreach ($User in $list){

    #set the pictue for the user
    SET-ADUser $User –add @{thumbnailphoto=$Picture}

    #output the user's name to the output file so we know who was changed
    $User.samaccountname >> $filename
}