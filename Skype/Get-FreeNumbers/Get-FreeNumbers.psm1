# get-freeNumbers
# queries AD and SFB for used extentions and will return a free extention that can be assigned to a new user
# Nathan Kewley 16/12/16
# ver 0.1

function global:get-FreeNumbers{
    PARAM(
        [parameter(HelpMessage='Add this parameter to check a specific number')]
        [string]$check
    )

    #connect to server
    import-module Remote-Hack 3>$null
    Remote-Hack SFB

    #create an array to hold all taken numbers
    [System.Collections.ArrayList]$takenNumbers = @{}

    #Try to populate all the used numbers
    write-host "Populating all used numbers..." -foregroundcolor Magenta
    try{
        #get all skype numbers
        write-host "Retrieving numbers from skype..." -foregroundcolor Magenta
        $takenNumbers += Get-CSuser -WarningAction SilentlyContinue | Select-Object LineURI
        $takenNumbers += Get-CsTrustedApplicationEndPoint -WarningAction SilentlyContinue | Select-Object LineURI
        $takenNumbers += Get-CsRgsWorkflow -Identity service:ApplicationServer:sr-vm-sfb-01.penrith.local -WarningAction SilentlyContinue | Select-Object LineURI 

        #get ad numbers
        write-host "Retrieving numbers from AD..." -foregroundcolor Magenta
        $ADNumbers = Get-ADUser -SearchBase “OU=New,dc=Penrith,dc=local” -Filter {telephonenumber -like "*"} -ResultSetSize 5000 -properties telephonenumber
    }catch{
        #throw an error if all numbers could not be retrieved
        throw "Could not get all used numbers"
        exit
    }

    #populate hard coded career limiting numbers
    write-host "Populating reserved numbers..." -ForegroundColor Magenta
    $clm = @(8399..8500)

    #add each of the numbers from AD
    write-host "Calculting AD Numbers..." -foregroundcolor Magenta
    foreach($AD in $ADNumbers){
        $ext = $AD.telephonenumber.substring($AD.telephonenumber.length - 4, 4)
        $clm += $ext
    }

    #for each clm we create a lineuri and add it to our array... maybe change to a linked list later for efficiancy
    Foreach($number in $clm){
        $number = 'tel:+6124732' + $number + ';ext=' + $number
        $clmNumber = New-Object System.Object
        $clmNumber | Add-Member -type NoteProperty -name LineURI -value $number
        $takenNumbers += $clmNumber | Select-Object LineURi
    }

    #if the user wants a ranom number we should find the next avaliable free number
    if(-not($check)){
        #find first avaliable free number
        write-host "Finding free number..." -ForegroundColor Magenta
        for($ext=7400; $ext -lt 8699; $ext++){
            #build the URI for the number
            $extURI = 'tel:+6124732' + $ext + ';ext=' + $ext
            $taken = $false
    
            #check the number against the lineURI's
            foreach($num in $takenNumbers){
                if($num.LineURI -eq $extURI){
                    $taken = $true
                }
            }

            #terminage loop if the URI is not taken and can be used
            if($taken -eq $false){break}
            #if($taken -eq $false){write-host $extURI -ForegroundColor green}
        }
    }

    #if the user weants to check a number avaliability we should do taht
    else{
        #find first avaliable free number
        write-host "Checking If Number Is Avaliable..." -ForegroundColor Magenta
        
        #build out lineURI       
        $taken = $false;
        $extURI = 'tel:+6124732' + $check + ';ext=' + $check

        #check the number against the lineURI's
        foreach($num in $takenNumbers){
            if($num.LineURI -eq $extURI){
                $taken = $true
            }
        }

        #terminage loop if the URI is not taken and can be used
        if($taken -eq $True){write-host "Number is taken"}
        else{write-host "Number is Free"}
    }

    #return selected free number
    return $extURI
}





