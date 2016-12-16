#ends a user's azure remote app session.
#Version 0.1
#Nathan Kewley
#14/11/2016

function global:end-azureSession{
    Param(
        [parameter(HelpMessage='your username/email for azure')]
        [string]$userEmail,
        [parameter(HelpMessage='your password for azure')]
        [string]$cred
    )

    #define our set variables for the penrith enviroment
    $azureSubscription = 'PCC-AAE'

    #create credentials
    #$secPass = ConvertTo-SecureString $password -AsPlainText -Force
    if($cred){
        #$AzureCred = New-Object System.Management.Automation.PSCredential ($username, $secPass)
    }else{
        $cred = Get-Credential
        #$AzureCred = New-Object System.Management.Automation.PSCredential
    }
    #this will throw exception even if success.... so yep... this is my work-around...
    try{Add-AzureAccount -Credential $cred}catch{write-host "Connected to Azure" -ForegroundColor green}

    #Select subscription
    Select-AzureSubscription $azureSubscription

    #disconnect user
    write-host "Disconnecting user, this may take a while....." -ForegroundColor green
    try{
        invoke-AzureRemoteAppSessionLogoff -CollectionName rappaaeprod -UserUpn $userEmail -confirm:$false >$null 2>&1
        write-host "diconnected..." -ForegroundColor green
    }catch{
        $errorMessage = $_.Exception.Message
        if($errorMessage -eq "InternalError: The server encountered an internal error. Please retry the request."){
            write-host "User has been disconnected" -ForegroundColor green
        }else{
            write-host "ERROR: User connection to Azure not found" -ForegroundColor red
        }
    }
}

