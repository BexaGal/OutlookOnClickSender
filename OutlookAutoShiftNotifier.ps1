Add-Type -Path $env:WINDIR\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Outlook.dll #API of Outlook
#It is shit that these DLLs contained in .Net Framework, not in .NET Core. Fuck...

if (!(test-path ./outlooksendermailconf.hash)){                  # First we should find out if config file exists.
    Write-Host "No config created. Creating one."               # if it deosn't, we create it in the current directory (where the file located).
    New-Item ./outlooksendermailconf.hash 
    $coalias = Read-Host "Write addressee alias"                # get alias of addressee from user
    $coaddress = Read-Host "Write E-mail address of $coalias"   # and it's E-mail
@"
[ordered]@{
    "$coalias"="$coaddress"
}
"@ | out-file ./outlooksendermailconf.hash
Write-Host "FOR ADDING MORE ENTRIES PLEASE EDIT THE CONFIG MANUALLY"
# here we just created first entry of the config file. Config contains raw hashtable.
}

if (!(test-path ./msgs.hash)){
    Write-Host "No messages file found. Creating one."
    New-Item ./msgs.hash
    $tempv1 = Read-Host -prompt "Write message's name"
    $tempb2 = Read-Host -prompt "Write message body"
@"
[ordered]@{
    "$tempv1"="$tempb2"
}
"@ | out-file ./msgs.hash
Write-Host "FOR ADDING MORE ENTRIES PLEASE EDIT THE CONFIG MANUALLY"
}

$hashcnfg = (Get-Content .\outlooksendermailconf.hash | Out-String)     # Here we utilise this config file, extracting the hashtable to a variable
$hashcnfg = ( Invoke-Expression $hashcnfg )                             # Why the fck we even need to do that?! Why it just can't get hashtable straight from a file???

$messagedata = (Get-Content .\msgs.hash | Out-String)
$messagedata = ( Invoke-Expression $messagedata )                        

$Outlook = New-Object -comobject Outlook.Application                                                # create an outlook instance
$namespace = $Outlook.GetNameSpace("MAPI")                                                          # MAPI namespace is used only for user's E-Mail extraction
$ebox = ((($namespace.Accounts | Select-Object {$_.DisplayName}) | ConvertTo-Csv)[1]).Trim('"')     # E-Mail extraction

Write-Host MAILBOX: $ebox                                                                           # Why tho... Well, in case if user's Outlook have more than one mailbox
                                                                                                    # So user will see that letter will be sent form wrong box. Ih he/she is not blind, ofc

Write-Host "Addressees:"                            # Getting addressees' mailboxes
$hashcnfg | Format-Table -Wrap

if ($hashcnfg.Count -ne 1){                            # Check if there is only one pair in hashtable 
    $taread = Read-Host -prompt "Choose addressee"     # If not, then user chooses addressee
    $adrread = $hashcnfg.$taread
}
else {
    $adrread = $hashcnfg[0]                             # Elseway programm does it itself
    Write-Host "Autosend to: " $adrread ". Add entries to outlooksendermailconf.hash to have more recipients"
}

Write-host "Here is messages' list:"
$messagedata | ft -wrap

Write-Host "Choose your destiny"                                                # Choosing template to apply
$destiny = Read-Host -prompt "Select message alias"
if ($messagedata.keys -contains $destiny){
    $template = $messagedata.$destiny
    Write-host "Message:"
}else {
    Write-Host "No such message. Abort."
    Start-sleep -seconds 3
    exit
}


$datet = Get-date -Format "dd.MM"                   # Getting current date
$Submess = "Смена, $datet"                          # We conseal it in the variable, containing subject

$message = $Outlook.CreateItem(0)                   # Creating letter instance
Write-Host "Sending letter to: "$adrread
$message.To = "$adrread"                              # filling addresee, subject and body of the letter
$message.Subject = $Submess
$message.Body = $template
$message.Send()                                     # Sending to addressee
Write-Host "Letter has been sent"                   # I guess a man, who is able to read, will not raise questions here
Start-Sleep -seconds 3
