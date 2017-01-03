#get all the users
$users = Get-ADuser -Filter * -SearchBase 'ou=UM, ou=test, ou=accounts, ou=new, dc=penrith, dc=local' -Properties *

#for each user get and change thier email address
forEach($user in $users){
    #get the email from ad
    $email = $user.EmailAddress

    #lowercase the email
    $emailLower = $email.ToLower();

    #update the users email address
    SET-ADUser $User -EmailAddress $emailLower

    #output changes
    write-host $email " -> " $emailLower;
}