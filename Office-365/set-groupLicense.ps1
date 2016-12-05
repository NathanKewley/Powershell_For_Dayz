##############################################################
# 365 AD group license                         Nathan Kewley #
# Powershell Version 4.0                  Ver 0.1 26/05/2016 #
#                                                            #
# Pulls all user from a specified AD group and assignes them #
# the specified license if they do no already have it        #
#                                                            #
# Usage                                                      #
# grantGroupLicense <AD-Group-Name> <License SkuPartNumber>  #
#                                                            #
# eg:                                                        #
# grantGroupLicense FG-MSVisio_Users VISIOCLIENT             #
# grantGroupLicense FG-MSProject_Users PROJECTCLIENT         #
##############################################################

function grantGroupLicense($group, $license){
    $members = Get-ADGroupMember $group -Recursive
    
    foreach($member in $members){
        $adUser = get-aduser $member.samaccountname | select userprincipalname
        $msolUser = Get-MsolUser -UserPrincipalName $adUser.userPrincipalName
        grantLicense $msolUser $license
    }

    function checkLicense($user, $license){
        foreach($lic in $user.licenses){
            if($lic.AccountSku.SkuPartNumber -eq $license){return $true}
        }
        return $false
    }

    function grantLicense($msolUser, $license){
        if(-not(checkLicense $msolUser $license)){
            try{
                set-msolUserLicense -userprincipalname $msolUser.userPrincipalName -AddLicenses $license  -ErrorAction Stop
                write-host $msolUser.userPrincipalName Has been given $license -ForegroundColor green
            }catch{
                write-host $msolUser.userPrincipalName could not be given $license -ForegroundColor red
            }
        }else{
            write-host $msolUser.userPrincipalName already has $license -ForegroundColor green
        }
    }
}
