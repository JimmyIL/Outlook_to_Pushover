do { 

Start-Sleep 10  


$outlook = New-Object -ComObject outlook.application
$namespace = $outlook.GetNameSpace("MAPI")
$folder=$namespace.GetDefaultFolder(6)
$unreadcount = $folder.UnReadItemCount

if ($unreadcount  -ge 1) { 
    
    $sendpush = @(
        $UserKey = "<youruserkeyhere>"
        $Message = "New Outlook Email"
        $device = "<yourphoneordevicenamehere>"
        $ApiKey = "<yourApiKeyhere>"
        $data = @{
           token = "$ApiKey";
            user = "$UserKey";
            message = "$Message";
            device = "$device" ;
        } 
    Invoke-RestMethod -Method Post -Uri "https://api.pushover.net/1/messages.json" -Body $data 
    )
    $sendpush 

    Start-Sleep 900
 }

}while($unreadcount -eq 0)
