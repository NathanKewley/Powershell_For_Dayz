#copnnect to mysol for online exchange
Connect-MsolService

$outputLocal = "C:\Users\P102542\Desktop\csv\ListofMailboxesLocal.csv"
$outputAzure = "C:\Users\P102542\Desktop\csv\ListofMailboxesAzure.csv"

#get all the on prem users
$users=Get-ADuser -Filter * -SearchBase 'ou=New, ou=User, ou=accounts, ou=new, dc=penrith, dc=local'

#for all of the on prem users
forEach($user in $users){
    $store = $user.GivenName + ',' + $user.Surname + ',' + $user.samaccountname + ',' + $user.userPrincipalName
    write-host $store
    $store >> $outputLocal
}

#get all of the licensed mysol users
$usersAZ = Get-MsolUser -MaxResults 2000 | Where-Object { $_.isLicensed -eq "TRUE" }
#$usersAZ = Get-MsolUser -MaxResults 2000

#loop thorugh each user
foreach ($user in $usersAZ) {
    $store = $user.DisplayName + ',' + $user.userPrincipalName
    write-host $store
    $store >> $outputAzure
}