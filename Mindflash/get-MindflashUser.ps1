##########################################################################################################
# get-MindflashUser cmdlet                                                            Ver 0.5 11/12/2015 #
# Powershell Version 4.0                                                                                 #
# If this script works it was written by Nathan Kewley                                                   #
# If this script does not work is was written by Michael Cho                                             #
# Get-MindflashUser is licensed under a Creative Commons Attribution 4.0 International License           #
#                                                                                                        #
# retrieves a mindflash user using the API by either email or SamAccountName, Displays the information   #
# about the user. password for both mindflash and ad can be reset using this script as well. You will    #
# need to pass in your api key to every command via the -key parameter or the script will fail           #
#                                                                                                        #
# ver 0.2                                                                                                #
# Added ability to get all users, and to get course info                                                 #
#                                                                                                        #
# ver 0.3                                                                                                #
# Wrapped script as a cmdlet among various improvements (run script once to include it in your poweshell)#
#                                                                                                        #
# ver 0.4                                                                                                #
# Added ability to invite users to courses, added ad flag for password resets to specify if as password  #
# should also be reset                                                                                   #
#                                                                                                        #
# ver 0.5                                                                                                #
# added ability to create new users. Password resets can now be 'only ad', 'only mindflash' or 'both'    #
#                                                                                                        #
# This script is intended to be used from the command line with parameters passed in. or can be used as  #
# a cmdlet from the command line/any script. Examples as follows:                                        #
#                                                                                                        #
# Get all users in mindflash                                                                             #
#   - .\get-MindflashUser                                                                                #
#                                                                                                        #
# Get all courses in mindflash                                                                           #
#   - .\get-MindflashUser -course all                                                                    #
#                                                                                                        #
# Get all users in a specific course                                                                     #
#   - .\get-MindflashUser -course 1369376710                                                             #
#                                                                                                        #
# Get all user information including all courses they are enrolled in                                    #
#   - .\get-MindflashUser -email user@email.com -course all                                              # 
#                                                                                                        #
# Get enrollemnt status for specific user in specific course including course progress                   #
#   - .\get-MindflashUser -email user@email.com -course 1333209300                                       #
#                                                                                                        #
# Get specific mindflash user by email:                                                                  #
#   - .\get-MindflashUser -email user@email.com                                                          #
#                                                                                                        # 
# Invite user to a specified course                                                                      #
#   - .\get-MindflashUser -course 1369376710 -sam p102542 -invite true                                   #
#                                                                                                        #
# resetting a users password for mindflash and ad:                                                       #
#    - .\get-MindflashUser -email user@email.com -resetPassword true -newPassword P@ssword21             #
#                                                                                                        #
# Creating a new user                                                                                    #
#    - .\get-MindflashUser -email user@email.com -first michael -last cho -newUser true -newPassword pass#
#                                                                                                        #
##########################################################################################################

#define the function that will act as the cmdlet
function global:get-MindflashUser{
    <#
        .SYNOPSIS
        provides easy admin for mindflash accoutns
        .DESCRIPTION
        retrieves a mindflash user using the API by either email or SamAccountName, Displays the information about the user. password for both mindflash and ad can  be reset using this script as well.
        .EXAMPLE                                             
        get-MindflashUser -email user@email.com -resetPassword true -newPassword P@ssword21
        Will reset the specified users password in both mindflash and ad
        .EXAMPLE                                             
        get-MindflashUser -course 1369376710
        Get all users in specific course
        .EXAMPLE
        get-MindflashUser -email user@email.com -course 1333209300
        Get enrollemnt status for specific user in specific course including course progress
    #>

    #add cmdlet binding
    [CmdletBinding()]

    #define required/optional parameters
    Param(
      [parameter(HelpMessage='Type the email of the user you want to modify')]
      [string]$email,
      [parameter(HelpMessage='Type the UPN of the user you want to modify (note, only enter an email or upn, not both')]
      [string]$sam,
      [parameter(HelpMessage='if set to true you must set the resetPassword parameter to the users new password')]
      [string]$resetPassword="false",
      [parameter(HelpMessage='what the new password for both mindflash and ad should be')]
      [string]$newPassword,
      [parameter(HelpMessage='flag to state if password should also be reset on ad, or just mindflash')]
      [string]$resetAD="false",
      [parameter(HelpMessage='Specify the AD OU that the user lives in')]
      [string]$ou,
      [parameter(HelpMessage='specify the api key, will default to our one')]
      [string]$key,
      [parameter(HelpMessage='specify a course id or put all for all courses')]
      [string]$course,
      [parameter(HelpMessage='specify if the user should be invited to the specified course')]
      [string]$invite="false",
      [parameter(HelpMessage='wether a new user should be created')]
      [string]$newUser="false",
      [parameter(HelpMessage='first name of new user')]
      [string]$first="false",
      [parameter(HelpMessage='last name of new user')]
      [string]$last="false"
    )
    Begin{}
    Process{

        #define some dirty globals
        $api = "https://api.mindflash.com/api/v2/"
        $mindflashID = $NULL
        $user = $NULL

        #make sure a key was entered, throw exception if not
        if(-not($PSBoundParameters.ContainsKey('key'))) {
            #error and terminate script
            $error = New-Object System.FormatException "No API Key Was Supplied, Terminating..."
            Throw $error
        }

        #if an email was entered
        if($PSBoundParameters.ContainsKey('email')) {
            #if upn was also entered parameters are not valid
            if ($PSBoundParameters.ContainsKey('sam')) {
                #error and terminate script
                $error = New-Object System.FormatException "Input Not Valid, set only email or sam. Not Both"
                Throw $error
            }

            #add the identification email to the api url
            $api = $api + "auth?email=$email"
        }

        #if a upn was entered
        if($PSBoundParameters.ContainsKey('sam')) {
            #we must query AD to get thier email for the mindflash api url
            $user = Get-ADUser -SearchBase "$ou" -LDAPFilter "(SamAccountName=$sam)" -Properties *

            #check if we were able to retrieve the user from ad
            if($user -ne $null){
                $email = $user.EmailAddress
            }
            #if cant retrieve user give error and terminate script
            else{
                $error = New-Object System.FormatException "User with sam: "$sam" could not be found in AD... Terminating"
                Throw $error
            }

            #add the identification email to the api url
            $api = $api + "auth?email=$email"
        }

        #if the request to create a new user has been made
        if($newUser -eq "true"){
            #check that required parameters have also been passed else throw an exception
            if(($PSBoundParameters.ContainsKey('newPassword')) -and ($PSBoundParameters.ContainsKey('first')) -and ($PSBoundParameters.ContainsKey('last')) -and ($PSBoundParameters.ContainsKey('email'))){
                #if requirements are met... request a new user be made
                write-host "Creating user: $email" -ForegroundColor Magenta
                
                #add date stamp
                $date = "{0:yyyy-mm-dd}" -f (get-date)

                #create post fields for api call
                $body = '
                {
                  "users": [
                    {
                      "firstName": "' + $first + '",
                      "lastName": "' + $last + '",
                      "email": "' + $email + '",
                      "password": "' + $newPassword + '"
                    }
                  ],
                  "clientDatestamp": "' + $date + '"
                }'

                #build api string
                $api = "https://api.mindflash.com/api/v2/user/"

                #try to create the user
                try{
                    #Invoke-RestMethod -Method POST -Uri $api -Body (ConvertTo-Json $body) -ContentType 'application/json' -Header @{ "X-mindflash-apiKey" = $Key }
                    Invoke-RestMethod -Method POST -Uri $api -Body $body -ContentType 'application/json' -Header @{ "X-mindflash-apiKey" = $Key }
                    write-host "mindflash user $email has been created" -ForegroundColor Green
                }catch{
                    $error = New-Object System.FormatException "User could not be created, please check input paremeters: 'first' 'last' 'newPassword' and 'email' "
                    Throw $error
                }
            }else{
                #error and throw a exception
                $error = New-Object System.FormatException "to create a new user you must specify parameters 'first' 'last' 'newPassword' and 'email'"
                Throw $error
            }
        }

        #if the reset password is set to true
        if($resetPassword -eq "true"){
            #make sure a new password is supplied
            if(-not($PSBoundParameters.ContainsKey('newPassword'))){
                #error and terminate script
                $error = New-Object System.FormatException "To reset user password you must provide a new password via the 'newPassword' parameter"
                Throw $error
            }

            #get the users sam addres from thier email if they are not searching by sam and ad password is to be reset
            if((-not($PSBoundParameters.ContainsKey('sam'))) -and ($resetAD -ne "false")){
                #query ad for the user based on thier email
                $user = Get-ADUser -SearchBase $ou -Filter {Emailaddress -eq $email}
    
                #check if we were able to retrieve the user from ad
                if($user -ne $null){
                    $sam = $user.SamAccountName
                }
                #if cant retrieve user give error and terminate script
                else{
                    $error = New-Object System.FormatException "User with Email Address: "$email" could not be found in AD... Terminating"
                    Throw $error
                }
            }

            #----reset mindflash passwd if required----#
            if($resetAD -ne "only"){
                #try get the user information from mindflash
                try{
                    $mindflashUser = Invoke-RestMethod -Method Get -Uri $api -Header @{ "X-mindflash-apiKey" = $Key }
                    $mindflashID = $mindflashUser.userId
                }catch{
                    $error = New-Object System.FormatException "Could not get mindflash user ID, please check input"
                    Throw $error
                }

                #let user know that mindflash password is being reset and build out api call string
                write-host "Resetting mindflash Password for user $mindflashID | $email" -ForegroundColor Magenta
                $body = @{
                    password=$newPassword
                    email=$email
                }
                $api = "https://api.mindflash.com/api/v2/user/" + $mindflashID
    
                #try reset the password
                try{
                    Invoke-RestMethod -Method POST -Uri $api -Body (ConvertTo-Json $body) -ContentType 'application/json' -Header @{ "X-mindflash-apiKey" = $Key }
                    write-host "mindflash password for user $email has been reset" -ForegroundColor Green
                }catch{
                    $error = New-Object System.FormatException "Check Password: Must contain at least 8 characters, one uppercase and lowercase letter, number, and symbol"
                    Throw $error
                }
            }

            #----reset AD passwd if required----#
            if($resetAD -eq "true"){
                #output that ad password is being reset
                write-host "Resetting AD Password for user $email" -ForegroundColor Magenta

                #get the ad user again because powershell does not seem to retain the objects returned for whatever silly reason it has to do that
                $user = Get-ADUser -SearchBase $ou -Filter {SamAccountName -eq $sam} -Properties *
    
                #check if a user was returned from ad. error out if not
                if($user -eq $NULL){
                    $error = New-Object System.FormatException "could not get user from AD... Terminating"
                    Throw $error
                }else{
                    #reset the users password... commented out for now
                    Set-ADAccountPassword $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $newPassword -Force)
                }

                write-host "AD Password for $email has been reset" -ForegroundColor Green
            }
        }

        #if a course is specified and no user return all users in the specified course
        if((-not($PSBoundParameters.ContainsKey('sam'))) -and (-not($PSBoundParameters.ContainsKey('email'))) -and ($PSBoundParameters.ContainsKey('course')) -and ($course -ne "all")){
            $api = $api+"course/"+$course+"/user"
        }

        #if a user is specified and course is set to all, all the courses the user is enrolled in should be queried
        if(($PSBoundParameters.ContainsKey('sam')) -or ($PSBoundParameters.ContainsKey('email'))){
            if(($PSBoundParameters.ContainsKey('course')) -and ($course -eq "all") -and ($invite -ne "true")){
                #try get the user information from mindflash
                try{
                    $mindflashUser = Invoke-RestMethod -Method Get -Uri $api -Header @{ "X-mindflash-apiKey" = $Key }
                    $mindflashID = $mindflashUser.userId
                }catch{
                    $error = New-Object System.FormatException "Could not get mindflash user ID, please check input"
                    Throw $error
                }
    
                $api = "https://api.mindflash.com/api/v2/user/" + $mindflashID

                #make the call
                Invoke-RestMethod -Method Get -Uri $api -Header @{ "X-mindflash-apiKey" = $Key }
            }
        }

        #if both a user and course are specified detailed information about the user enrolment in the course should be returned
        if(($PSBoundParameters.ContainsKey('sam')) -or ($PSBoundParameters.ContainsKey('email'))){
            if(($PSBoundParameters.ContainsKey('course')) -and ($course -ne "all") -and ($invite -ne "true")){
                #try get the user information from mindflash... this script is such a mess now but oh well, it works
                try{
                    $mindflashUser = Invoke-RestMethod -Method Get -Uri $api -Header @{ "X-mindflash-apiKey" = $Key }
                    $mindflashID = $mindflashUser.userId
                }catch{
                    $error = New-Object System.FormatException "Could not get mindflash user ID, please check input"
                    Throw $error
                }
    
                $api = "https://api.mindflash.com/api/v2/course/" + $course + "/user/" + $mindflashID

                #make the call
                Invoke-RestMethod -Method Get -Uri $api -Header @{ "X-mindflash-apiKey" = $Key }
            }
        }

        #if inviting user to a course
        if($invite -eq "true"){
            if(($PSBoundParameters.ContainsKey('sam')) -or ($PSBoundParameters.ContainsKey('email'))){
                #check that a single course id has been entered
                if(($PSBoundParameters.ContainsKey('course')) -and ($course -ne "all")){
                    #try get the user information from mindflash... this script is such a mess now but oh well, it works
                    try{
                        $mindflashUser = Invoke-RestMethod -Method Get -Uri $api -Header @{ "X-mindflash-apiKey" = $Key }
                        $mindflashID = $mindflashUser.userId
                    }catch{
                        $error = New-Object System.FormatException "Could not get mindflash user ID, please check input"
                        Throw $error
                    }
                
                    #if user id valid build the command                
                    $api = "https://api.mindflash.com/api/v2/course/" + $course + "/user/" + $mindflashID+"/invite"

                    $date = "{0:yyyy-mm-dd}" -f (get-date)
                    $body = @{
                        clientDatestamp=$date
                        required="false"
                    }
                    
                    #make the call
                    try{ 
                        Invoke-RestMethod -Method POST -Uri $api -Body (ConvertTo-Json $body) -ContentType 'application/json' -Header @{ "X-mindflash-apiKey" = $Key }
                        write-host "user $mindflashID has been invited to course $course" -ForegroundColor green
                    }catch{
                        #error and terminate script
                        $error = New-Object System.FormatException "Could not add user to specified course, check course is correct and is not archived"
                        Throw $error
                    }
                }else{
                    #error and terminate script
                    $error = New-Object System.FormatException "please make sure a single course id is specified"
                    Throw $error
                }
            }else{
                #error and terminate script
                $error = New-Object System.FormatException "You must specify a user to enroll"
                Throw $error
            }
        }

        #display the info if the user is just after info
        if(($mindflashID -eq $NULL) -and ($newUser -eq "false") -and ($resetAD -ne "only")){
            try{
                #if no user or course parameter is supplied we should return all users
                if((-not($PSBoundParameters.ContainsKey('sam'))) -and (-not($PSBoundParameters.ContainsKey('email'))) -and (-not($PSBoundParameters.ContainsKey('course')))){$api = $api+"user/"}

                #if no users are supplied and all courses are requested show all courses
                if((-not($PSBoundParameters.ContainsKey('sam'))) -and (-not($PSBoundParameters.ContainsKey('email'))) -and ($course -eq "all")){$api = $api+"course/"}

                #make the call
                Invoke-RestMethod -Method Get -Uri $api -Header @{ "X-mindflash-apiKey" = $Key }
            }catch{
                #if command failed output error and command that was attempted
                write-host "Invoke-RestMethod threw exception EX: " $_.Exception.Message " | " $api -ForegroundColor red
            }
        }

        #return true
        return $True
    }
    End{}
}