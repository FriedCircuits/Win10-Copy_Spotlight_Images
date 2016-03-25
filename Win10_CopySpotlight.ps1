#Windows 10 Spotlight Background Copier
#Copies Spotlight image to desktop and renames to .jpg
#Author: William Garrido
#Created: 03-07-2016
#Updated: 03-11-2016 Removes files that are not widescreen
$version = "1.0"
#License: CC-BY-SA
Write-Host "Windows 10 Spotlight Image Copy - Version:"$version
$spotpath = $env:USERPROFILE + '\AppData\Local\Packages\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\LocalState\Assets'
$dest = $env:USERPROFILE + '\Desktop\Spotlight'
$files = Get-ChildItem $spotpath
echo "Scanning for new images to copy..."
if(!(Test-Path -Path $dest)){
    New-Item -ItemType directory -Path $dest
}
$count = 0
Foreach ($file in $files){
    $copyfile = "false"
    $check = $dest + "\" + $file + ".jpg" 
    if(!(Test-Path -Path $check) -and $file.Length -gt 400KB){
        Copy-Item $spotpath\$file $dest -PassThru | 
        Rename-Item  -NewName {$_.Name + ".jpg"} 
        $copyfile = "true"
        $count++
    }
    if($copyfile -eq "true"){
       $newFile = Get-Item $check
       $shellObject = New-Object -ComObject Shell.Application
       $directoryObject = $shellObject.NameSpace($newfile.Directory.FullName)
       $fileObject = $directoryObject.ParseName($newfile.Name)
       $dimString = $directoryObject.GetDetailsOf($fileObject, 31)
       #echo $dimString
       if($dimString -eq "1080 x 1920"){ Remove-Item $newFile; $count--}
    }
}

if($count -eq 0){echo "No new images."} else {Write-Host "New Images:" $count}