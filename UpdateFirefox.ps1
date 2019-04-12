#Created to check and download latest FireFox
#Checks Live Version vs what is deployed, then downloads and email notification if newer version is found
#Creation date : 4-2-2019
#Creator: Alix N Hoover



#Variables to your software Share
$SCCMSource = '\\sccm\Software\Mozilla'

#Variables-Mail
$MailServer = "MAIL"
$recip = "TO"
$sender = "FROM"
$subject = "Firefox Update"
#where you put your schedule task
$ServerName = "SCCM" 
#where you put your documentation on deploying files
$doc='onenote:///L:\Documentation\---04CB35}&page-id={BB768083-B904-48AF-BB7D-32DA865C8BAE}&end'



#weblinks
$uricheck = 'https://product-details.mozilla.org/1.0/firefox_versions.json'
$URIx86 = 'https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win&lang=en-US'
$URI = 'https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=en-US'


#Check Current Version
$FirefoxVersion = Invoke-WebRequest $uricheck -UseBasicParsing | ConvertFrom-Json | select -ExpandProperty LATEST_FIREFOX_vERSION
$FirefoxVersion
$checkfolder = "$SCCMSource\$FirefoxVersion"

IF (!(test-path $checkfolder)) {
    Write-Output "$checkfolder does not exist Proceeding With Download"

#Variables-64
$OutFile    = 'Firefox Setup x64.msi'
$OutFile = "$SCCMSource\$OutFile"

#Variables-86
$OutFilex86 = 'Firefox Setup x86.msi'
$OutFileX86 = "$SCCMSource\$OutFileX86"



# Download FireFox from the web
Write-Output "Downloading $URI to $OutFile"
Invoke-WebRequest -Uri $URI -OutFile $OutFile -UserAgent [Microsoft.PowerShell.Commands.PSUserAgent]::explorer
Write-Output "Downloading $URIx86 to $OutFilex86"
Invoke-WebRequest -Uri $URIx86 -OutFile $OutFilex86 -UserAgent [Microsoft.PowerShell.Commands.PSUserAgent]::explorer

# Get file metadata 64bit

$a = 0 
$objShell = New-Object -ComObject Shell.Application 
$objFolder = $objShell.namespace((Get-Item $OutFile).DirectoryName) 

foreach ($File in $objFolder.items()) {
    IF ($file.path -eq $outfile) {
        $FileMetaData = New-Object PSOBJECT 
        for ($a ; $a  -le 266; $a++) {  
         if($objFolder.getDetailsOf($File, $a)) { 
             $hash += @{$($objFolder.getDetailsOf($objFolder.items, $a)) = $($objFolder.getDetailsOf($File, $a)) }
            $FileMetaData | Add-Member $hash 
            $hash.clear()  
           } #end if 
       } #end for  
    }
}

# Move the downloaded file to the appropriate location
$Version = $FileMetaData.subject
$Version = $Version.split(' ')[2]

Write-Output "Downloaded version: $Version"
$Filename = $FileMetaData.subject
$Filename = $Filename + ".msi"
$destinationfolder = "$SCCMSource\$Version"
Write-Output "Destination folder is $destinationfolder"


   [System.IO.Directory]::CreateDirectory($destinationfolder)
  Write-Output "Creating $destinationfolder"
    [System.IO.File]::Move($OutFile,"$destinationfolder\$Filename")
   Write-Output "Moving $OutFile to $destinationfolder"



# Get file metadata 32bit

$a = 0 
$objShellx86 = New-Object -ComObject Shell.Application 
$objFolderx86 = $objShellx86.namespace((Get-Item $OutFilex86).DirectoryName) 

foreach ($Filex86 in $objFolderx86.items()) {
    IF ($filex86.path -eq $outfilex86) {
        $FileMetaDatax86 = New-Object PSOBJECT 
        for ($a ; $a  -le 266; $a++) {  
         if($objFolderx86.getDetailsOf($Filex86, $a)) { 
             $hashx86 += @{$($objFolderx86.getDetailsOf($objFolderx86.items, $a)) = $($objFolderx86.getDetailsOf($Filex86, $a)) }
            $FileMetaDatax86 | Add-Member $hashx86 
            $hashx86.clear()  
           } #end if 
       } #end for  
    }
}

# Move the downloaded file to the appropriate location
$Versionx86 = $FileMetaDatax86.subject
$Versionx86 = $Versionx86.split(' ')[2]
$Filenamex86 = $FileMetaDatax86.subject
$Filenamex86 = $Filenamex86 + ".msi"
[System.IO.File]::Move($OutFilex86,"$destinationfolder\$Filenamex86")
   Write-Output "Moving $OutFilex86 to $destinationfolder"

#send email
$body ="<html></body> <BR> FireFox Version <p style='color:#FF0000'>  $Version </p> is ready for deployment Via SCCM  <BR>"
$body+= "<a href=$doc>Here are Directions</a> "
$body+="<BR> this is a Scheduled task on $ServerName"
 
Send-MailMessage -From $sender -To $recip -Subject $subject -Body ( $Body | out-string ) -BodyAsHtml -SmtpServer $MailServer



}
   
   ELSE { Write-Output "$checkfolder already exists"} 
