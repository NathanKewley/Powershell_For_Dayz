##############################################################
# PowerShell Enterprise Vault items finder        08/09/2015 #
# Powershell Version 4.0                                     #
# If this script works it was written by Nathan Kewley       #
# If this script does not work is was written by Michael Cho #
#                                                            #
# Scan thorugh a user's mailbox and retrieve all of thier    #
# vaulted items while maintaining folder structure           #
#                                                            #
# NOTE: User must have whole mailbox cached offline          #
##############################################################

#file path saving output
$emailPathLog = "c:\Scans\emailFiles.csv";

#initial script setup
$ol = new-object -com Outlook.Application
$ns = $ol.GetNamespace("MAPI")

#variables to keep track of current folder to keep folder structure in tact
$masterDir = "c:\vaultedmail\" + $ol.Application.DefaultProfileName  + "\";
$currentDir = "c:\vaultedmail\" + $ol.Application.DefaultProfileName  + "\";

#counter so emails are not overwritten
$counter = 0;

#get the vault folder
$vault = $namespace.Folders | ?{$_.name -match "Inbox"}

#check if the email is vaulted, if so move it to the local drive keeping folder structure in tact
function checkMail($mail){
    if($mail.body -like "*has been archived*"){
        write-host("Vaulted Item Detected" + $counter);
        $filename = $currentDir + "mail" + $counter + ".msg";
        $mail.SaveAs($filename);
        $filename >> $emailPathLog;
    }
}

#recursivly scan all items and subfolders in users mailbox
function Get-MailboxFolder($folder){
    write-host $folder.name, $folder.items.count

    #create directory for the folder
    $path = $currentDir + $folder.name + "\";
    New-Item -ItemType directory -Path $path;
    $currentDir = $path;

    foreach($mail in $folder.items){
        $counter++;
        checkMail($mail);
    }

    foreach ($f in $folder.folders){
        Get-MailboxFolder $f
    }
}

#start a transcript
Start-Transcript -Path "c:\Scans\log.txt";

#initial script setup
$ol = new-object -com Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$mailbox = $ns.stores | where {$_.ExchangeStoreType -eq 0}
$mailbox.GetRootFolder().folders | foreach { Get-MailboxFolder $_}