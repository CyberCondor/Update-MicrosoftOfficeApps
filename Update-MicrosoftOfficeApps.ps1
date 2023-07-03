$OfficeC2RClientPath = "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"
if(Test-Path $OfficeC2RClientPath){
   if(((Get-Item $OfficeC2RClientPath).VersionInfo.ProductName     -eq "Microsoft*Office") -and 
      ((Get-Item $OfficeC2RClientPath).VersionInfo.FileDescription -eq "Microsoft Office Click-to-Run Client")){
       cmd /c '"C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" /update user updatepromptuser=true forceappshutdown=false displaylevel=true'
   }
}
